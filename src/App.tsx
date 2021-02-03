import React from 'react';
import logo from './logo.svg';
import './App.css';
import { Authentication, PowerAppsConnection } from './Authentication/Authentication';
import { AuthParams } from './RunSettings.development';
import { DefaultButton } from "@fluentui/react";
export default class App extends React.Component<{}, {}> {

  async authenticate() {
    debugger;
    let response: PowerAppsConnection = await Authentication.authenticate(AuthParams);
    console.log(response);
  }

  render() {
    return (
      <DefaultButton text="Standard" onClick={this.authenticate} allowDisabledFocus />
    )
  }
}


