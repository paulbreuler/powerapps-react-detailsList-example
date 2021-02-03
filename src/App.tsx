import React from 'react';
import logo from './logo.svg';
import './App.css';
import { Authentication, PowerAppsConnection } from './Authentication/Authentication';
import { AuthParams, EnvironmentDetails } from './RunSettings.development';
import { DefaultButton, DetailsList, Fabric, IColumn, TextField } from "@fluentui/react";

interface IAppState {
  powerAppsConnection: PowerAppsConnection,
  listItems: any[],
  listColumns: IColumn[],
  query: string
}

export default class App extends React.Component<{}, IAppState> {
  constructor(props: any) {
    super(props);
    this.authenticate();
    this.state = {
      ...this.state,
      listItems: [],
      listColumns: [],
      query: "accounts"
    }
  }

  async authenticate(): Promise<void> {
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
        columnActionsMode: 1,
        onColumnClick: this._onColumnClick,
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

    this.setState({ ...this.state, listItems: mappedResponse, listColumns: columns })

  }

  private onChangeText = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: string | undefined): void => {
    this.setState({ ...this.state, query: text as string })
  };


  private _onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    debugger;
    const { listColumns, listItems: listCollection } = this.state;
    const newColumns: IColumn[] = listColumns.slice();
    const currColumn: IColumn = newColumns.filter(currCol => column.key === currCol.key)[0];
    newColumns.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });
    const newItems = this._copyAndSort(listCollection, currColumn.fieldName!, currColumn.isSortedDescending);
    this.setState({
      listColumns: newColumns,
      listItems: newItems,
    });
  }

  private _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
    const key = columnKey as keyof T;
    return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
  }

    render() {
      return (
        <Fabric>
          <TextField value={this.state.query} label="Api Query" onChange={this.onChangeText} />
          <DefaultButton text="Submit Request" onClick={() => this.apiRequest(this.state.query)} allowDisabledFocus />
          { this.state.listItems.length > 0 && this.state.listColumns.length > 0 ? <DetailsList columns={this.state.listColumns} items={this.state.listItems} /> : null}
        </Fabric>
      )
    }
  }


