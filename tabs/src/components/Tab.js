import './App.css';
import './Tab.css';
import React from 'react';
import { TeamsFx, createMicrosoftGraphClient } from '@microsoft/teamsfx';
import * as microsoftTeams from '@microsoft/teams-js';
import { Button, Dropdown, Input, Flex, Card, Avatar, Text, Grid } from "@fluentui/react-northstar"
import {
    SearchIcon, BriefcaseIcon, EmailIcon, AppFolderIcon, CallIcon, ChevronStartIcon,
    FilterIcon
} from "@fluentui/react-icons-northstar";

import { Profile, } from "./Profile";
import defaultPhoto from '../images/default-photo.png';
import appLogo from './about/images/logo_2_30_x_30.png';


class Tab extends React.Component {

    constructor(props) {
        super(props);
        this.state = {
            userInfo: {},
            showLoginPage: undefined,
            allUsers: [],
            departments: [],
            selectedDepartment: "",
            searchText: "",
            filteredUsers: [],
            searchedUsers: [],
            selectedUser: "",
            onSmallScreen: undefined,
            showProfileListing: undefined,
            showProfileDisplay: undefined,
            isConfigured: undefined,
            platformInfo: {},
        }
    }

    async componentDidMount() {
        await this.initTeamsSDK();
        await this.initTeamsFx();
        await this.initData();
        console.log(this.state.allUsers);
    }

    async initTeamsSDK() {
        try {
            await microsoftTeams.app.initialize();
            const context = await microsoftTeams.app.getContext();

            if (Object.values(microsoftTeams.HostName).includes(context.app.host.name)) {
                microsoftTeams.app.notifySuccess();
            }
        } catch (error) {
            microsoftTeams.app.notifyFailure(
                {
                    reason: microsoftTeams.app.FailedReason.Timeout,
                    message: error
                }
            )
        }
    }

    async initTeamsFx() {
        const teamsfx = new TeamsFx();
        // Get the user info from access token and set tenant's configured state
        let userInfo;
        try {
            userInfo = await teamsfx.getUserInfo();

            this.setState({
                userInfo: userInfo,
                isConfigured: true
            });
        } catch (error) {
            // This is used to catch the error that occurs when app is deployed
            // on a tenant that's not configured yet
            if (error.message?.includes("resourceDisabled")) {
                this.setState({
                    isConfigured: false
                });
            }
        }

        this.teamsfx = teamsfx;
        this.scope = ["User.Read", "User.ReadBasic.All", "User.Read.All", "Contacts.Read"];
        const platformInfo = await this.getPlatformInfo();
        this.setState({
            platformInfo: platformInfo,
        })
    }

    async initData() {
        if (!await this.checkIsConsentNeeded()) {
            // Initialize graph client
            const graphClient = createMicrosoftGraphClient(this.teamsfx, this.scope);
            this.graphClient = graphClient;

            // Check user's device screen size
            await this.checkScreenCategory();

            // fetch all users without profile photo
            await this.getAllUsers();

            // set default selected user to currently logged-in user
            this.setState({
                selectedUser: this.state.userInfo.displayName
            });

            // extract department for dropdown
            await this.setDepartments();

            // fetch all users' profile photo
            await this.getUsersPhotos();
        }
    }

    async loginBtnClick() {
        try {
            // Popup login page to get user's access token
            await this.teamsfx.login(this.scope);
        } catch (err) {
            if (err instanceof Error && err.message?.includes("CancelledByUser")) {
                const helpLink = "https://aka.ms/teamsfx-auth-code-flow";
                err.message +=
                    "\nIf you see \"AADSTS50011: The reply URL specified in the request does not match the reply URLs configured for the application\" " +
                    "in the popup window, you may be using unmatched version for TeamsFx SDK (version >= 0.5.0) and Teams Toolkit (version < 3.3.0) or " +
                    `cli (version < 0.11.0). Please refer to the help link for how to fix the issue: ${helpLink}`;
            }

            alert("Login failed: " + err);
            return;
        }
        await this.initData();
    }

    async checkIsConsentNeeded() {
        try {
            await this.teamsfx.getCredential().getToken(this.scope);
        } catch (error) {
            this.setState({
                showLoginPage: true
            });
            return true;
        }
        this.setState({
            showLoginPage: false
        });
        return false;
    }


    async getPlatformInfo() {
        return new Promise((resolve, reject) => {
            microsoftTeams.app.getContext().then((context) => {
                const name = context.app.host?.name;
                resolve({ name: name });
            })
        })
    }

    async getAllUsers() {
        try {
            await this.graphClient
                .api("/users?$top=999")
                .filter("(onPremisesSyncEnabled eq true OR userType eq 'Member') and accountEnabled eq true")
                .select(["id", "mail", "displayName", "jobTitle", "mail", "mobilePhone", "department",
                    "userPrincipalName", "businessPhones", "employeeId", "userType", "accountEnabled", "onPremisesSyncEnabled"])
                .get(async (error, response) => {
                    if (!error) {
                        this.setState({
                            allUsers: response.value.sort((a,b) => (a.displayName > b.displayName) ? 1 : ((b.displayName > a.displayName) ? -1 : 0))
                        })
                        
                        Promise.resolve();
                    } else {
                        console.log("graph error", error);
                    }
                });
        } catch (error) {
            console.log(error);
        }
    }

    async getUsersPhotos() {
        const usersWithPhoto = await Promise.all(
            this.state.allUsers.map(async (user) => {
                try {
                    const response = await this.graphClient
                        .api(`/users/${user.id}/photo/$value`)
                        .get()
                    user.profilePhoto = URL.createObjectURL(response);
                    return user
                } catch (error) {
                    if (error.statusCode === 404) {
                        user.profilePhoto = "";
                        return user;
                    } else {
                        console.log(error);
                    }
                }
            })
        );

        this.setState({
            allUsers: usersWithPhoto
        })
    }

    async setDepartments() {
        let departments = this.state.allUsers
            .filter((user) => Boolean(user.department))
            .map((user) => user.department.trim());
        let uniqueDepartments = [...new Set(departments)];
        uniqueDepartments.unshift("All");
        this.setState({
            departments: uniqueDepartments
        });
    }


    async setFilteredUsers(selectedItem) {

        if (selectedItem === "All") {
            this.setState({
                filteredUsers: this.state.allUsers
            });
        } else {
            let filteredUsers = this.state.allUsers.filter((user) => user.department === selectedItem);
            this.setState({
                filteredUsers: filteredUsers
            })
        }
    }

    async setSearchedUsers(searchText) {
        // the check to see if filteredUsers is empty is used because the dropdown has no default
        // selection which makes the filteredUsers list empty when the page is newly loaded
        if (searchText !== "" && this.state.filteredUsers.length === 0) {
            let searchedUsers = this.state.allUsers.filter((user) => {
                return user.displayName.toLowerCase().includes(searchText.toLowerCase());
            });
            this.setState({
                searchedUsers: searchedUsers
            });
        } else if (searchText !== "" && this.state.filteredUsers.length > 0) {
            let searchedUsers = this.state.filteredUsers.filter((user) => {
                return user.displayName.toLowerCase().includes(searchText.toLowerCase());
            });
            this.setState({
                searchedUsers: searchedUsers
            });
        } else {
            // TODO: check how this block relates to the conditional rendering of users list
            // and see if probably you can optimize the conditions in the conditional rendering
            this.setState({
                searchedUsers: this.state.filteredUsers
            })
        }
    }



    async checkScreenCategory() {
        const layout = document.getElementsByClassName('layout')[0];
        const styles = window.getComputedStyle(layout);
        const marginBottom = styles.marginBottom;

        if (marginBottom === "1px") {
            this.setState({
                onSmallScreen: true,
                showProfileListing: true,
                showProfileDisplay: false
            });
        } else {
            this.setState({
                onSmallScreen: false,
                showProfileListing: true,
                showProfileDisplay: true
            });
        }
    }



    async handleVisibility() {
        const filterListBox = document.getElementsByClassName('filter-list-box')[0];
        const actionText = document.getElementsByClassName('action-button')[0];

        const filterListBoxStyles = window.getComputedStyle(filterListBox);

        if (filterListBoxStyles.display === 'none') {
            filterListBox.style.display = 'block';
            actionText.style.color = '#a9a9a9'
        } else {
            filterListBox.style.display = 'none';
            actionText.style.color = '';
        }
    }

    render() {
        // Functions for rendering the list of users within an organization
        // Function 1: DOM
        let usersListDom = (usersList) => {
            return usersList.map((user) => {
                return (
                    <Card
                        id={user.displayName}
                        aria-roledescription="card avatar"
                        centered size="small"
                        onClick={(_, event) => {
                            if (this.state.onSmallScreen) {
                                this.setState({
                                    selectedUser: event.id,
                                    showProfileListing: false,
                                    showProfileDisplay: true
                                });
                            } else {
                                this.setState({
                                    selectedUser: event.id
                                });
                            }
                        }}
                        className='card-dimensions'
                    // styles={{ padding: '10px 0 0 0', marginBottom: '9px' }}
                    >
                        <Card.Header>
                            <Flex gap="gap.smaller" column hAlign="center">
                                <Avatar
                                    image={user.profilePhoto ? user.profilePhoto : defaultPhoto}
                                    label=""
                                    name=""
                                    size="larger"
                                />
                                <Flex column hAlign="center">
                                    <Text content={user.displayName} weight="bold" align="center" styles={{ margin: "0px 5px" }} />
                                    <Text content={user.jobTitle ? user.jobTitle : "N/A"} size="small" />
                                </Flex>
                            </Flex>
                        </Card.Header>
                    </Card>
                )
            })
        };

        // Function 2: Conditional rendering
        let usersList = () => {
            if (this.state.searchText === "" && this.state.filteredUsers.length === 0) {
                return usersListDom(this.state.allUsers);
            } else if (this.state.searchText === "" && this.state.filteredUsers.length > 0) {
                return usersListDom(this.state.filteredUsers);
            } else {
                return usersListDom(this.state.searchedUsers);
            }
        }


        // Functions for displaying the full details of a user
        // Function (for mobile) : Handle back button onclick
        const returnButton = () => {
            this.setState({
                showProfileListing: true,
                showProfileDisplay: false,
                searchText: "" // this is added to reset search results if one has been made prior to viewing profile
            });
        }


        // Function : Handle opening of Microsoft Teams chat window
        const handleChatOpening = async (user) => {
            const chatParams = {
                user: user.userPrincipalName || '',
            };
            await microsoftTeams.chat.openChat(chatParams);
        }

        // Function : Handle opening of Microsoft Teams chat window
        const handleAudioCall = async (user) => {
            const callParams = {
                targets: [`${user.userPrincipalName}`],
            };
            await microsoftTeams.call.startCall(callParams);
        }

        // Function 1: DOM
        let userProfileDom = (user) => {
            return (
                <div className='display-wrapper'>
                    {this.state.onSmallScreen === true && <div className='buttons-container'>
                        <button className='display-back-button' onClick={returnButton}><ChevronStartIcon />Back</button>
                        {microsoftTeams.chat.isSupported() === true && <div className='contact-button'>
                            <button className='contact-button-chat' onClick={() => handleChatOpening(user)}>Chat</button>
                            <button onClick={() => handleAudioCall(user)}>Call</button>
                        </div>}
                    </div>}

                    <div className="display-profile">
                        <div >
                            {this.state.onSmallScreen === false && microsoftTeams.chat.isSupported() === true && <div className='contact-button'>
                                <button className='contact-button-chat' onClick={() => handleChatOpening(user)}>Chat</button>
                                <button onClick={() => handleAudioCall(user)}>Call</button>
                            </div>
                            }
                        </div>
                        <Flex column hAlign="center">
                            <div className="photo">
                                <img src={user.profilePhoto ? user.profilePhoto : defaultPhoto} alt="avatar" />
                            </div>
                            <Text content={user.displayName ? user.displayName : "N/A"} weight="bold" size="large" />
                        </Flex>

                        <Flex column gap="gap.large" styles={{ width: "85%", margin: "10px auto 0px" }}>
                            <p style={{ margin: "10px 5px" }} ><EmailIcon /><span></span>{user.mail ? user.mail : "N/A"}</p>
                            <p style={{ margin: "10px 5px" }} ><BriefcaseIcon /><span></span>{user.jobTitle ? user.jobTitle : "N/A"}</p>
                            <p style={{ margin: "10px 5px" }} ><AppFolderIcon /><span></span>{user.department ? user.department : "N/A"}</p>
                            <p style={{ margin: "10px 5px" }} ><CallIcon /><span></span>{user.mobilePhone ? user.mobilePhone : "N/A"}</p>
                        </Flex>
                    </div>
                </div>
            )
        }

        // Function 2: Conditional rendering
        let userProfile = () => {
            if (this.state.selectedUser) {
                let user = this.state.allUsers.find((user) => user.displayName === this.state.selectedUser);
                return userProfileDom(user);
            }
        }


        // Functions for handling profile listing filtering
        // Function 1: Filter listings based on selected item
        let filterListings = async (e) => {
            // Get selected list item text
            const selectedInnerHTML = e.target.innerHTML;

            changeItemFontWeight(selectedInnerHTML);
            this.setState({
                selectedDepartment: selectedInnerHTML
            });
            await this.setFilteredUsers(selectedInnerHTML);
            await this.setSearchedUsers(this.state.searchText);

            // Hide filter list box and change filter action text color
            const filterListBox = document.getElementsByClassName('filter-list-box')[0];
            const actionText = document.getElementsByClassName('action-button')[0];

            filterListBox.style.display = 'none';
            actionText.style.color = '';
        }

        // Function 2: Change selected list item's font weight
        let changeItemFontWeight = (selectedInnerHTML) => {
            const allItems = document.getElementsByClassName('filter-item');
            for (let i = 0; i < allItems.length; i++) {
                if (allItems[i].innerHTML === selectedInnerHTML) {
                    allItems[i].style.fontWeight = "bold";
                } else {
                    allItems[i].style.fontWeight = "";
                }
            }
        }

        // Function 3: Populate list items
        let listItems = () => {
            return this.state.departments.map((item) => {
                return (
                    <li className='filter-item' onClick={filterListings}>{item}</li>
                )
            })
        }


        return (
            <div className='tab-page'>
                {this.state.isConfigured === true && <div>
                    {this.state.showLoginPage === false &&
                        <div className='layout'>

                            {/* section 1 */}
                            {this.state.showProfileListing && <div className='listing'>

                                {/* Listing header for mobile view */}
                                <div className='mobile-listing-header'>
                                    <div className='search-input-container'>
                                        <input type="text" placeholder='Search employee...'
                                            onInput={(event) => {
                                                let searchText = event.target.value;
                                                //console.log(searchText);
                                                this.setState({
                                                    searchText: searchText
                                                })
                                                this.setSearchedUsers(searchText);
                                            }}
                                        />
                                    </div>

                                    <div className='filter-container'>
                                        <div className='filter-action-text-box'>
                                            <button className='action-button' onClick={async () => { await this.handleVisibility() }}><FilterIcon /> Click to filter by department</button>
                                        </div>
                                        <div className='filter-list-box'>
                                            <ul>
                                                {listItems()}
                                            </ul>
                                        </div>
                                    </div>
                                </div>

                                {/* Listing header for desktop view */}

                                <div className='listing-header-wrapper'>
                                    <div className='listing-header'>
                                        <div className='listing-header-dropdown'>
                                            <Dropdown
                                                items={this.state.departments}
                                                placeholder="Filter by department"
                                                checkable
                                                getA11ySelectionMessage={{
                                                    onAdd: item => `${item} has been selected.`,
                                                }}
                                                fluid
                                                onChange={async (_, event) => {
                                                    this.setState({
                                                        selectedDepartment: event.value
                                                    });
                                                    await this.setFilteredUsers(event.value);
                                                    await this.setSearchedUsers(this.state.searchText);
                                                }}
                                            />
                                        </div>

                                        <div className='listing-header-search'>
                                            <Input
                                                icon={<SearchIcon />}
                                                placeholder="Search Employee..."
                                                fluid
                                                onChange={(_, event) => {
                                                    this.setState({
                                                        searchText: event.value
                                                    })
                                                    this.setSearchedUsers(event.value);
                                                }}
                                            />
                                        </div>
                                    </div>
                                </div>

                                {/* Profile listing */}
                                <div className='profile-cards-container'>
                                    {this.state.onSmallScreen === false && <Flex gap="gap.smaller" wrap="true" hAlign="center">
                                        {usersList()}
                                    </Flex>}
                                    {this.state.onSmallScreen === true && <div className='listing-grid-mobile'>
                                        {usersList()}
                                    </div>}
                                </div>
                            </div>}


                            {/* Section 2 */}
                            {this.state.showProfileDisplay === true && <div className='display'>
                                {userProfile()}
                            </div>}
                        </div>
                    }

                    {this.state.showLoginPage === true && <div className="auth">
                        <Profile userInfo={this.state.userInfo} />
                        <h2>The Employee Lookup Application</h2>
                        <Button primary onClick={() => this.loginBtnClick()}>Start</Button>
                    </div>}
                </div>}

                {this.state.isConfigured === false &&
                    <div className='install-error-wrapper'>
                        <div className='install-error'>
                            <div><img src={appLogo} alt='Employee Lookup Logo' className='logo' /></div>
                            <p className='install-error-head'>Almost there! Just a couple configurations required.</p>
                            <p className='install-error-body'>You are seeing this page because your tenant is not properly configured to run the Employee Lookup application.</p>
                            <div className='error-decision'>
                                <div>
                                    <p>Please contact support using this email</p>
                                    <p className='contact'>be@relianceinfosystems.com</p>
                                </div>
                                <div> Or</div>
                                <div>
                                    <p>Visit our product page to learn more</p>
                                    <div><a className='contact link' href="{{state.fx-resource-frontend-hosting.endpoint}}" target="_blank" rel="noopener noreferrer">Click to visit our product page</a></div>
                                </div>
                            </div>
                        </div>
                    </div>
                }
            </div>

        )
    }
}

export default Tab;