import React from 'react';
import logo from './logo.svg';
import './App.css';
import { Authentication, PowerAppsConnection } from './Authentication/Authentication';
import { AuthParams, EnvironmentDetails } from './RunSettings.development';
import { DefaultButton } from "@fluentui/react";

interface IAppState {
  powerAppsConnection: PowerAppsConnection
}

export default class App extends React.Component<{}, IAppState> {
  constructor(props: any) {
    super(props);
    this.authenticate();
  }

  async authenticate() {
    debugger;
    let response: PowerAppsConnection = await Authentication.authenticate(AuthParams);
    this.setState({ ...this.state, powerAppsConnection: response })
  }

  async apiRequest(query: string) {
    let response = await fetch(`${EnvironmentDetails.org_url}/${query}`,
      {
        method: "GET", headers: {
          accept: "application/json",
          "OData-MaxVersion": "4.0",
          "OData-Version": "4.0",
          "Content-Type": "application/json; charset=utf-8",
          Authorization: `Bearer ${this.state.powerAppsConnection.access_token}`
        }
      });

    if (response.status !== 200) {
      console.log("Request Failed!");
    }

    let responseJson = await response.json();

    console.log(responseJson);

  }

  render() {
    return (
      <DefaultButton text="Standard" onClick={() => this.apiRequest("accounts")} allowDisabledFocus />
    )
  }
}


