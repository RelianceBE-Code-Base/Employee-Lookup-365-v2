import './App.css';
import './Tab.css';
import React from 'react';
import { TeamsFx } from '@microsoft/teamsfx';
import * as microsoftTeams from '@microsoft/teams-js';

import TabImplentation from './Tab';
import appLogo from './about/images/logo_2_30_x_30.png';


class Tab extends React.Component {

    constructor(props) {
        super(props);
        this.state = {
            isConfigured: undefined
        }
    }

    async componentDidMount() {
        await this.initTeamsSDK();

        if (await this.checkTenantConfigStatus()) {
            this.setState({ isConfigured: true })
        } else {
            this.setState({ isConfigured: false })
        }
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

    async checkTenantConfigStatus() {
        try {
            let teamsfx = new TeamsFx();
            await teamsfx.getUserInfo();
            return true
        } catch (err) {
            if (err.message?.includes("resourceDisabled")) {
                return false
            }

        }
    }

    render() {
        return (
            <div className='tab-page'>
                {this.state.isConfigured === true && <div>
                    <TabImplentation />
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