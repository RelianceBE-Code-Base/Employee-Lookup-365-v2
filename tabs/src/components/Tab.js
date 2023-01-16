import './App.css';
import './Tab.css';
import React from 'react';
import { TeamsFx, BearerTokenAuthProvider, createApiClient } from '@microsoft/teamsfx';
import * as microsoftTeams from '@microsoft/teams-js';
import { Dropdown, Input, Flex, Card, Avatar, Text } from "@fluentui/react-northstar"
import {
    SearchIcon, BriefcaseIcon, EmailIcon, AppFolderIcon, CallIcon, ChevronStartIcon,
    FilterIcon
} from "@fluentui/react-icons-northstar";
import { Buffer } from "buffer"

// import { Profile, } from "./Profile";
import defaultPhoto from '../images/default-photo.png';
import appLogo from './about/images/logo_2_30_x_30.png';
import config from '../lib/config';

class TabImplentation extends React.Component {

    constructor(props) {
        super(props);
        this.state = {
            userInfo: {},
            showConsentPage: false,
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
            consentErrorText: "",
            consentErrorHeader: "",
            hostName: ""
        }
    }

    async componentDidMount() {
        await this.initTeamsFx();
        await this.initData();
    }

    async initTeamsFx() {
        const teamsfx = new TeamsFx();
        // Get the user info from access token
        let userInfo;
        try {
            userInfo = await teamsfx.getUserInfo();
            this.setState({ userInfo: userInfo });
        } catch (err) {
            console.log(err);
        }

        this.teamsfx = teamsfx;
        this.scope = ["User.Read", "User.ReadBasic.All", "User.Read.All"];
    }

    async initData() {
        try {
            // Initialize Azure Functions API Client
            const credential = this.teamsfx.getCredential();
            const apiBaseUrl = config.apiEndpoint + "/api/";
            this.apiClient = createApiClient(
                apiBaseUrl,
                new BearerTokenAuthProvider(async () => (await credential.getToken("")).token)
            );

            // Check user's device screen size
            await this.checkScreenCategory();

            // set host name
            await this.getHostName();

            // get all users without profile photo from azure functions
            await this.getAllUsers();

            // set default selected user to currently logged in user
            this.setState({ selectedUser: this.state.userInfo.objectId });

            // extract department for dropdown
            await this.setDepartments();

            // get all users with their photo from azure functions
            await this.getUsersPhotos();
        } catch (err) {
            if (err.message.includes("invalid_grant")) {
                this.setState({ showConsentPage: true })
            } else if (err.message.includes("Unable to get")) {
                alert(err.message);
            } else {
                // alert("An error occured!");
                console.log(err);
                alert(err);
            }
        }
    }

    async loginBtnClick() {
        try {
            // Popup login page to get user's access token
            await this.teamsfx.login(this.scope);
            this.setState({ showConsentPage: false });
            await this.initData();
        } catch (err) {
            let message;
            if (err instanceof Error && (err.message?.includes("CancelledByUser") || err.message?.includes("User declined"))) {
                message = "The consent process was cancelled. Please grant consent or contact your admin in order to use the application.";
                this.setState({ consentErrorHeader: "Consent Cancelled!" })
                this.setState({ consentErrorText: message })
            } else if (err instanceof Error && err.message?.includes("browser is blocking the url to open")) {
                message = "The consent process was blocked. If you are an admin, kindly consent in Microsoft Teams and refresh the application here.";
                this.setState({ consentErrorHeader: "Consent Blocked!" })
                this.setState({ consentErrorText: message })
            } else {
                alert("An error occured!");
            }
        }
    }

    async getHostName() {
        let context = await microsoftTeams.app.getContext();
        let hostName = context.app.host?.name;
        this.setState({ hostName: hostName });
        return hostName;
    }

    async getAllUsers() {
        try {
            const response = await this.apiClient.get("allUsers");
            this.setState({ allUsers: response.data.allUsers });
        } catch (err) {
            throw new Error(err?.response?.data?.error);
        }
    }

    async getUsersPhotos() {
        try {
            let response = await this.apiClient.get("allUsersWithPhoto");
            //convert JSON buffer to base64 image string
            response = response.data.allUsers.map((user) => {
                if (user.profilePhoto) {
                    let photo = Buffer(user.profilePhoto);
                    photo = photo.toString("base64");
                    user.profilePhoto = photo;
                };
                return user;
            })

            this.setState({ allUsers: response });
        } catch (err) {
            throw new Error(err.response.data.error);
        }
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
        let screenSize = window.innerWidth;

        if (screenSize <= 820) {
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

    async handleFilterVisibility() {
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
                        id={user.id}
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
                                    image={user.profilePhoto ? `data:image/jpeg;base64,${user.profilePhoto}` : defaultPhoto}
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

        // Function : Handle opening of Microsoft Teams call window
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
                    {this.state.onSmallScreen === true &&
                        <div className='buttons-container'>
                            <button className='display-back-button' onClick={returnButton}><ChevronStartIcon />Back</button>
                            {microsoftTeams.chat.isSupported() === true && this.state.selectedUser !== this.state.userInfo.objectId && <div className='contact-button'>
                                <button className='contact-button-chat' onClick={() => handleChatOpening(user)}>Chat</button>
                                <button onClick={() => handleAudioCall(user)}>Call</button>
                            </div>}
                        </div>
                    }

                    <div className="display-profile">
                        <div >
                            {this.state.onSmallScreen === false && microsoftTeams.chat.isSupported() === true && this.state.selectedUser !== this.state.userInfo.objectId &&
                                <div className='contact-button'>
                                    <button className='contact-button-chat' onClick={() => handleChatOpening(user)}>Chat</button>
                                    <button onClick={() => handleAudioCall(user)}>Call</button>
                                </div>
                            }
                        </div>
                        <Flex column hAlign="center">
                            <div className="photo">
                                <img src={user.profilePhoto ? `data:image/jpeg;base64,${user.profilePhoto}` : defaultPhoto} alt="avatar" />
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
                let user = this.state.allUsers.find((user) => user.id === this.state.selectedUser);
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
                {this.state.showConsentPage === false &&
                    <div className='layout'>

                        {/* section 1 */}
                        {this.state.showProfileListing && <div className='listing'>

                            {/* Listing header for mobile view */}
                            <div className='mobile-listing-header'>
                                <div className='search-input-container'>
                                    <input type="text" placeholder='Search employee...'
                                        onInput={(event) => {
                                            let searchText = event.target.value;
                                            this.setState({
                                                searchText: searchText
                                            })
                                            this.setSearchedUsers(searchText);
                                        }}
                                    />
                                </div>

                                <div className='filter-container'>
                                    <div className='filter-action-text-box'>
                                        <button className='action-button' onClick={async () => { await this.handleFilterVisibility() }}><FilterIcon /> Click to filter by department</button>
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

                {this.state.showConsentPage === true &&
                    <div className='install-error-wrapper'>
                        <div className='install-error'>
                            <div><img src={appLogo} alt='Employee Lookup Logo' className='logo' /></div>
                            <p className='install-error-head'>{(this.state.consentErrorHeader) ? this.state.consentErrorHeader : "Consent Required!"}</p>
                            <p className='install-error-body'>{(this.state.consentErrorText) ? this.state.consentErrorText : "To continue using the app, some permissions are needed. Kindly click the consent button below if you are an admin or contact an administrator."}</p>
                            <div className='consent-button-wrapper'>
                                <button className='consent-button' onClick={() => this.loginBtnClick()}>Consent</button>
                            </div>
                        </div>
                    </div>
                }
            </div>

        )
    }
}

export default TabImplentation;