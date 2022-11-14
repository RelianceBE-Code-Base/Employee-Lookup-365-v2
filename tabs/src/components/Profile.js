// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React from 'react';
import { /*Flex, Text, Flex, Header, List, Checkbox, Input, MenuButton*/ } from "@fluentui/react-northstar";
import './Profile.css';
import defaultPhoto from '../images/default-photo.png';
import { /*BriefcaseIcon, EmailIcon, AppFolderIcon, CallIcon,  PersonIcon */ } from "@fluentui/react-icons-northstar";

class Profile extends React.Component {
  render() {
    return (
      <div className="profile">
        <div className="photo">
          <img src={defaultPhoto} alt="avatar" />
        </div>
        <div className="info">
          <div className="name">{this.props.userInfo.displayName}</div>
          <div className="email">{this.props.userInfo.preferredUserName}</div>
        </div>
      </div>
    );
  }
}

class Profile2 extends React.Component {
  render() {
    let user = this.props.user;
    console.log("user props");
    console.log(user);
    // let userProfileDom = (user) => {
    //   return (
    //     <div style={{ margin: "0 20px" }}>
    //       <div className="profile2">
    //         <Flex column hAlign="center">
    //           <div className="photo">
    //             <img src={defaultPhoto} alt="avatar" />
    //           </div>
    //           <Text content={user.displayName} weight="bold" size="large" />
    //         </Flex>

    //         <Flex column gap="gap.large" styles={{ width: "85%", margin: "10px auto 0px" }}>
    //           <p style={{ margin: "10px 5px" }} ><EmailIcon /><span></span>{user.mail ? user.mail : "N/A"}</p>
    //           <p style={{ margin: "10px 5px" }} ><BriefcaseIcon /><span></span>{user.jobTitle ? user.jobTitle : "N/A"}</p>
    //           <p style={{ margin: "10px 5px" }} ><AppFolderIcon /><span></span>{user.department ? user.department : "N/A"}</p>
    //           <p style={{ margin: "10px 5px" }} ><CallIcon /><span></span>{user.mobilePhone ? user.mobilePhone : "N/A"}</p>
    //         </Flex>
    //       </div>
    //     </div>
    //   )
    // };

    // let userProfile = () => {
    //   if (this.props.id.selectedID) {
    //     let user = this.props.allUsers.find((user) => user.displayName === this.props.id.selectedID);
    //     return userProfileDom(user);
    //   } else {
    //     let user = this.props.allUsers.find((user) => user.displayName === this.props.userInfo.displayName);
    //     return userProfileDom(user);
    //   }
    // };

    return (
      // <div style={{ margin: "0 20px" }}>
      //   <div className="profile2">
      //     <Flex column hAlign="center">
      //       <div className="photo">
      //         <img src={defaultPhoto} alt="avatar" />
      //       </div>
      //       <Text content={user.displayName ? user.displayName : "N/A"} weight="bold" size="large" />
      //     </Flex>

      //     <Flex column hAlign="" gap="gap.large" styles={{ width: "85%", margin: "10px auto 0px" }}>
      //       <p style={{ margin: "10px 5px" }} ><EmailIcon /><span></span>{user.mail ? user.mail : "N/A"}</p>
      //       <p style={{ margin: "10px 5px" }} ><BriefcaseIcon /><span></span>{user.jobTitle ? user.jobTitle : "N/A"}</p>
      //       <p style={{ margin: "10px 5px" }} ><AppFolderIcon /><span></span>{user.department ? user.department : "N/A"}</p>
      //       <p style={{ margin: "10px 5px" }} ><CallIcon /><span></span>{user.mobilePhone ? user.mobilePhone : "N/A"}</p>
      //     </Flex>
      //   </div>
      // </div>

      <div>
      {/* {userProfile()}  */}
      </div>
    );
  }
}

export { Profile, Profile2 };
