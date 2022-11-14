// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React from 'react';
import '../App.css';

/**
 * This component is used to display the required
 * terms of use statement which can be found in a
 * link in the about tab.
 */
class Support extends React.Component {
    render() {
      return (
        <div>
          <h1>Support</h1>
          <p>To contact suuport, please use the following phone numbers</p>
          <ul>
            <li>BUSINESS SUPPORT: (+234) 1-628-2729</li>
            <li>CUSTOMER SERVICE: (+234) 1-628-2729</li>
          </ul>
        </div>
      );
    }
}

export default Support;