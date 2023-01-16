import React from "react";
import { Link } from 'react-router-dom';
import './css/style.css';
import './css/utilities.css';
import topImage from './images/top_image.png';
import desktopMain from './images/desktop_1_main.png';
import desktopFilter from './images/desktop_2_filter.png';
import desktopSearch from './images/desktop_3_search.png';
import desktopProfile from './images/desktop_4_profile.png';
import mobileMain from './images/mobile_1_main.png';
import mobileFilter from './images/mobile_2_filter.png';
import mobileSearch from './images/mobile_3_search.png';
import mobileProfile from './images/mobile_4_profile.png';

class Home extends React.Component {
    render() {
        return (
            <div className="homepage">

                {/* Navbar */}
                <div className="navbar">
                    <div className="container flex">
                        <h1 className="logo">Employee Lookup</h1>
                        <nav>
                            <ul>
                                <li><Link to="/">Home</Link></li>
                                <li><Link to="/privacy">Privacy</Link></li>
                                <li><Link to="/termsofuse">Terms of Use</Link></li>
                            </ul>
                        </nav>
                    </div>
                </div>

                {/* Showcase */}
                <section class="showcase">
                    <div class="container grid">

                        <div class="showcase-text">
                            <h1>Everyone</h1>
                            <h1>in your organization</h1>
                            <h1>at a glance.</h1>
                            <p>As organizations grow in size and scales, keeping track of all employees becomes hard. The distance and connection between employees grows wider. Employee Lookup brings all your employees together by giving them easy access to each others contact information.</p>
                            <p>Built for Microsoft Teams, Office 365 and Outlook.</p>
                            <a href="/" class="btn btn-primary">Check on AppSource</a>
                        </div>

                        <div class="showcase-image-container">
                            {/* <div class="desktop-image"><img src={desktopMain} alt="" /></div>
                            <div class="mobile-image"><img src={mobileMain} alt="" /></div> */}
                            <div class="top-image"><img src={topImage} alt="" /></div>
                        </div>

                    </div>
                </section>

                {/* App features */}
                <section className="app-features">
                    <div className="section-header">
                        <h1>The Application Features</h1>
                    </div>

                    <div className="app-features-everyone">
                        <div className="grid container">
                            <div>
                                <h2>See everyone</h2>
                                <p>When you login into the app, what you see is the listing of all employees and your profile sitting right next to it.</p>
                            </div>
                            <div><img src={desktopMain} alt="" /></div>
                        </div>
                    </div>

                    <div className="app-features-filter">
                        <div className="grid container">
                            <div><img src={desktopFilter} alt="" /></div>
                            <div>
                                <h2>Filter the list</h2>
                                <p>You can filter down the employee listing based on the
                                    departments available within your organization.</p>
                            </div>
                        </div>
                    </div>

                    <div className="app-features-search">
                        <div className="grid container">
                            <div>
                                <h2>Search for a colleague</h2>
                                <p>Using the search box at the top of the app, you can
                                    directly search for a colleague by entering an alphabet&#40;s&#41; or their name.</p>
                            </div>
                            <div><img src={desktopSearch} alt="" /></div>
                        </div>
                    </div>

                    <div className="app-features-profile">
                        <div className="grid container">
                            <div><img src={desktopProfile} alt="" /></div>
                            <div>
                                <h2>View profile</h2>
                                <p>Click on a profile card to view available details for the profile.</p>
                            </div>
                        </div>
                    </div>
                </section>

                {/* App features on mobile */}
                <section className="app-features-mobile">
                    <div className="section-header">
                        <h1>When you are on the go</h1>
                    </div>
                    <div className="grid container">
                        <div className="item item-1"><img src={mobileMain} alt="employee list screen for mobile" /></div>
                        <div className="item item-2"><img src={mobileFilter} alt="filter employee list by department for mobile" /></div>
                        <div className="item item-3"><img src={mobileSearch} alt="search employee by name for mobile" /></div>
                        <div className="item item-4"><img src={mobileProfile} alt="view a profile for mobile" /></div>
                    </div>
                </section>

                {/* The build */}
                <section className="the-build" id="the-make-up">
                    <div className="section-header">
                        <h1>The Make Up</h1>
                    </div>
                    <div className="grid container">
                        <div className="the-build-tools">
                            <h2>Azure Active Directory</h2>
                            <h2>Microsoft Graph</h2>
                            <h2>And more...</h2>
                        </div>
                        <div>
                            <p>The employee lookup application is built leveraging different Microsoft technologies to provide you with the experience of having access to everyone in your organization at one glance. These includes the Azure Active Directory and the Microsoft Graph.</p>
                            <p>To use the Employee Lookup application, your admin staff would create an Azure AD application, which would be connected to the application code to bring you this experience. This is to give secure access to the Office 365 users within your organization. </p>
                            <p>And we wouldn't just jump in, an admin would have to consent to the permissions before anyone within your organization can use the app.</p>
                            <p>We are here to make the configuration easy for you, please contact support at <span>be@relianceinfosystems.com</span></p>
                        </div>
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
        )
    }
}

export default Home;