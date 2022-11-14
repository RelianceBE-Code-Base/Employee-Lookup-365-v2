// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React from 'react';
import { Link } from 'react-router-dom';
import '../App.css';

/**
 * This component is used to display the required
 * terms of use statement which can be found in a
 * link in the about tab.
 */
class TermsOfUse extends React.Component {
  render() {
    return (
      <div className='terms-of-use'>
        {/* Navbar */}
        <div class="navbar">
          <div class="container flex">
            <h1 class="logo">Employee Lookup</h1>
            <nav>
              <ul>
                <li><Link to="/">Home</Link></li>
                <li><Link to="/privacy">Privacy</Link></li>
                <li><Link to="/termsofuse">Terms of Use</Link></li>
              </ul>
            </nav>
          </div>
        </div>

        {/* Privacy */}
        <section class="terms-of-use-text">
          <div class="container">
            <h1>Terms of Use</h1>
            <p>Last updated: October 03, 2022</p>

            <p>The Employee Lookup application is a copyrighted work belonging to Reliance Infosystems Limited. Certain
              features of the application may be subject to additional guidelines, terms, or rules, which will be
              posted here in connection with such features.</p>

            <p>All such additional terms, guidelines, and rules are incorporated by reference into these Terms of Use
              document.</p>

            <p>These Terms of Use described the legally binding terms and conditions that oversee your use of the
              Application. BY USING THE APPLICATION, YOU ARE BEING COMPLIANT TO THESE TERMS, and you represent that
              you have the authority and capacity to enter into these Terms. IF YOU DISAGREE WITH ALL OF THE PROVISION
              OF THESE TERMS, DO NOT USE THE APPLICATION.</p>


            <h2>Definition</h2>
            <p>Company &#40;referred to as either "the Company", "We", "Us" or "Our" in this Agreement&#41; refers to
              Reliance Infosystems Limited.</p>


            <h2>Access to the Application</h2>

            <p><strong>Subject to these Terms.</strong> Company grants you a non-transferable, non-exclusive, revocable,
              limited license to access the Site solely for your own personal, noncommercial use.</p>

            <p><strong>Certain Restrictions.</strong> The rights approved to you in these Terms are subject to the
              following restrictions: &#40;a&#41; you shall not sell, rent, lease, transfer, assign, distribute, host,
              or otherwise commercially exploit the application; &#40;b&#41; you shall not change, make derivative
              works of, disassemble, reverse compile or reverse engineer any part of the application; &#40;c&#41; you
              shall not access the application in order to build a similar or competitive application; and &#40;d&#41;
              except as expressly stated herein, no part of the application may be copied, reproduced, distributed,
              republished, downloaded, displayed, posted or transmitted in any form or by any means unless otherwise
              indicated, any future release, update, or other addition to functionality of the application shall be
              subject to these Terms. All copyright and other proprietary notices on the application must be retained
              on all copies thereof.</p>

            <p>Company reserves the right to change, suspend, or cease the application with or without notice to you.
              You approved that Company will not be held liable to you or any third-party for any change,
              interruption, or termination of the application or any part.</p>

            <p>Excluding any User Content that you may provide, you are aware that all the intellectual property rights,
              including copyrights, patents, trademarks, and trade secrets, in the application and its content are
              owned by Company or Company&#39;s suppliers. Note that these Terms and access to the application do not
              give you any rights, title or interest in or to any intellectual property rights, except for the limited
              access rights expressed above. Company and its suppliers reserve all rights not granted in these Terms.
            </p>


            <h2>Limitation on Liability</h2>

            <p>To the maximum extent permitted by law, in no event shall company or our suppliers be liable to you or
              any third-party for any lost profits, lost data, costs of procurement of substitute products, or any
              indirect, consequential, exemplary, incidental, special or punitive damages arising from or relating to
              these terms or your use of, or incapability to use the application even if company has been advised of
              the possibility of such damages. Access to and use of the application is at your own discretion and
              risk, and you will be solely responsible for any damage to your device or computer system, or loss of
              data resulting therefrom.</p>


            <h2>Copyright Policy</h2>

            <p>Company respects the intellectual property of others and asks that users of our application do the same.
              In connection with our application, we have adopted and implemented a policy respecting copyright law
              that provides for the removal of any infringing materials and for the termination of users of our
              application who are repeated infringers of intellectual property rights, including copyrights.</p>


            <h2>General</h2>

            <p><strong>Third-Party Sites.</strong> The application may contain links to third-party websites. You should
              be aware that these third-party websites may collect, use or transfer your personal data in accordance
              with their personal data protection policies. And that we have no control over these websites, their
              content or the processing of your personal data by such websites. Company cannot be held liable in
              connection therewith.</p>

            <p><strong>Entire Terms.</strong> These Terms constitute the entire agreement between you and us regarding
              the use of the application. Our failure to exercise or enforce any right or provision of these Terms
              shall not operate as a waiver of such right or provision. The section titles in these Terms are for
              convenience only and have no legal or contractual effect. The word "including" means "including without
              limitation". If any provision of these Terms is held to be invalid or unenforceable, the other
              provisions of these Terms will be unimpaired and the invalid or unenforceable provision will be deemed
              modified so that it is valid and enforceable to the maximum extent permitted by law.</p>

            <p><strong>Your Privacy.</strong> Please read our Privacy Policy.</p>

            <p><strong>Copyright/Trademark Information.</strong> Copyright ©. All rights reserved. All trademarks, logos
              and service marks displayed on the Site are our property or the property of other third-parties. You are
              not permitted to use these Marks without our prior written consent or the consent of such third party
              which may own the Marks.</p>

            <h2>Contact Information</h2>
            <p>If you have any questions about this Terms of Use document, please contact support at <span>support@relianceinfosystems.com</span></p>
          </div>
        </section>

        {/* Footer */}
        <footer className="footer">
          <div>
            <p><Link to="/privacy">Privacy Policy</Link></p>
            <p><Link to="/termsofuse">Terms of Use</Link></p>
            <p>&copy; Reliance Infosystems Limited</p>
          </div>
        </footer>
      </div>
    );
  }
}

export default TermsOfUse;