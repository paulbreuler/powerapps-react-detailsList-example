import React from 'react';
import logo from './logo.svg';
import './App.css';
import { Authentication, PowerAppsConnection } from './Authentication/Authentication';
import { AuthParams, EnvironmentDetails } from './RunSettings.development';
import { DefaultButton, DetailsList, Fabric, IColumn, TextField } from "@fluentui/react";

interface IAppState {
  powerAppsConnection: PowerAppsConnection,
  listCollection: any[],
  listColumns: IColumn[],
  query: string
}

export default class App extends React.Component<{}, IAppState> {
  constructor(props: any) {
    super(props);
    this.authenticate();
    this.state = {
      ...this.state,
      listCollection: [],
      listColumns: [],
      query: "accounts"
    }
  }

  async authenticate(): Promise<void> {
    debugger;
    let response: PowerAppsConnection = await Authentication.authenticate(AuthParams);
    this.setState({ ...this.state, powerAppsConnection: response })
  }

  async apiRequest(query: string): Promise<void> {
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
    let idKey = "";
    if (this.state.query.endsWith("ies")) {
      idKey = `${this.state.query.substring(0, this.state.query.length - 3)}yid`
    } else {
      idKey = `${this.state.query.substring(0, this.state.query.length - 1)}id`
    }

    let columns: any = [];

    const addColumn = (col: any) => {
      columns.push({
        key: col,
        name: col,
        fieldName: col,
        minWidth: 35,
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
    };

    addColumn(idKey);

    // Get only keys with name to filter columns on view and allow dynamic object types.
    Object.keys(responseJson.value[0]).filter(item => { return item.includes('name') && !item.includes('yomi') && !item.includes('address') }).forEach((column: any) => {
      debugger;
      addColumn(column);
    });

    let mappedResponse: {}[] = [];

    // Map elements into a simplified array for use in DetailsList. 
    responseJson.value.forEach((elem: any) => {
      let newElem: any = {};
      newElem[idKey] = elem.opportunityid;
      columns.forEach((col: any) => {
        newElem[col.fieldName] = elem[col.fieldName]
      })
      mappedResponse.push(newElem)
    })

    this.setState({ ...this.state, listCollection: mappedResponse, listColumns: columns })

  }

  private onChangeText = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: string | undefined): void => {
    this.setState({ ...this.state, query: text as string })
  };

  render() {
    return (
      <Fabric>
        <TextField value={this.state.query} label="Api Query" onChange={this.onChangeText} />
        <DefaultButton text="Submit Request" onClick={() => this.apiRequest(this.state.query)} allowDisabledFocus />
        { this.state.listCollection.length > 0 && this.state.listColumns.length > 0 ? <DetailsList columns={this.state.listColumns} items={this.state.listCollection} /> : null}
      </Fabric>
    )
  }
}


