import React from 'react';
import logo from './logo.svg';
import './App.css';
import { Authentication, PowerAppsConnection } from './Authentication/Authentication';
import { AuthParams, EnvironmentDetails } from './RunSettings.development';
import { DefaultButton, DetailsList, Fabric, IColumn } from "@fluentui/react";

interface IAppState {
  powerAppsConnection: PowerAppsConnection,
  listCollection: any[],
  listColumns: IColumn[]
}

export default class App extends React.Component<{}, IAppState> {
  constructor(props: any) {
    super(props);
    this.authenticate();
    this.state = {
      ...this.state,
      listCollection:[],
      listColumns: [],
    }
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
    
    // To be more generic you would need to be able to accept interfaces that can then be deconstructed
    var mappedResponse = responseJson.value.map((item: any) => { const container = { "accountId": item.accountid, "name": item.name }; return container; })

    let columns: any = [];
    Object.keys(mappedResponse[0]).forEach((column: any) => {
      debugger;
      columns.push({
        key: column,
        name: column,
        fieldName: column,
        minWidth: 100,
        maxWidth: 200,
        isCollapsible: true,
        isCollapsable: true,
        isGrouped: false,
        isMultiline: false,
        isResizable: true,
        isRowHeader: false,
        isSorted: false,
        isSortedDescending: false,
        columnActionsMode: 1
      })
    });

    this.setState({ ...this.state, listCollection: mappedResponse, listColumns: columns })

  }

  render() {
    return (
      <Fabric>
        <DefaultButton text="Standard" onClick={() => this.apiRequest("accounts")} allowDisabledFocus />
        { this.state.listCollection.length > 0 && this.state.listColumns.length > 0 ? <DetailsList columns={this.state.listColumns} items={this.state.listCollection} /> : null}
      </Fabric>
    )
  }
}


