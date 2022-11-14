// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React from 'react';
import './App.css';
import * as microsoftTeams from "@microsoft/teams-js";

/**
 * The 'Config' component is used to display your group tabs
 * user configuration options.  Here you will allow the user to
 * make their choices and once they are done you will need to validate
 * their choices and communicate that to Teams to enable the save button.
 */
 class TabConfig extends React.Component {

  render() {
    // Initialize the Microsoft Teams SDK
    microsoftTeams.app.initialize();

    /**
     * When the user clicks "Save", save the url for your configured tab.
     * This allows for the addition of query string parameters based on
     * the settings selected by the user.
     */
    microsoftTeams.pages.config.registerOnSaveHandler((saveEvent) => {
      const baseUrl = `https://${window.location.hostname}:${window.location.port}`;
      microsoftTeams.pages.config.setConfig({
        "suggestedDisplayName": "Employee Lookup",
        "entityId": "Test",
        "contentUrl": baseUrl + "/index.html#/tab",
        "websiteUrl": baseUrl + "/index.html#/tab"
      });
      saveEvent.notifySuccess();
    });

    /**
     * After verifying that the settings for your tab are correctly
     * filled in by the user you need to set the state of the dialog
     * to be valid.  This will enable the save button in the configuration
     * dialog.
     */
    microsoftTeams.pages.config.setValidityState(true);

    return (
      <div>
        <h1>Tab Configuration</h1>
        <div style={{ marginLeft: '20px' }}>
          <p>The application doesn't need any special configuration.</p>
          <p>Please click the Save button to continue.</p>
        </div>
      </div>
    );
  }
}

export default TabConfig;