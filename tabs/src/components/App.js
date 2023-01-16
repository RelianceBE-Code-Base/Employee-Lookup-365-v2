// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React from 'react';
import './App.css';
import * as microsoftTeams from '@microsoft/teams-js';
import { HashRouter as Router, Route } from 'react-router-dom';

import Home from './about/homepage';
import Privacy from './about/Privacy';
import TermsOfUse from './about/TermsOfUse';
import Tab from './landing';
import TabConfig from './TabConfig';
import Support from './about/Support';

/**
 * The main app which handles the initialization and routing
 * of the app.
 */
function App() {
  // Check for the Microsoft Teams SDK object.
  if (microsoftTeams) {
    return (
      <Router>
        <Route exact path='/' component={Home} />
        <Route exact path='/tab' component={Tab} />
        <Route exact path='/config' component={TabConfig} />
        <Route exact path='/privacy' component={Privacy} />
        <Route exact path='/termsofuse' component={TermsOfUse} />
        <Route exact path='/support' component={Support} />
      </Router>
    );
  } else {
    return (
      <h3>Microsoft Teams SDK not found.</h3>
    );
  }
}

export default App;
