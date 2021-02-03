import React from 'react';
import './App.css';
import { Authentication, PowerAppsConnection } from './Authentication/Authentication';
import { AuthParams, EnvironmentDetails } from './RunSettings.development';
import { DefaultButton, DetailsList, DetailsListLayoutMode, Fabric, getId, IColumn, Stack, TextField } from "@fluentui/react";
import { Pagination } from "@uifabric/experiments";
import { buttonStyles, sectionStackTokens } from './AppStyles';

interface IAppState {
  powerAppsConnection: PowerAppsConnection,
  listItems: any[],
  listColumns: IColumn[],
  query: string,
  itemsPerPage: number,
  selectedPageIndex: number;
}

export default class App extends React.Component<{}, IAppState> {
  private _allItems: any[] = [];

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
    let idKey = this.getIdKey();

    let columns = this.loadColumns(idKey, responseJson);
    let mappedResponse = this.loadItems(idKey, columns, responseJson);
    this._allItems = mappedResponse;

    let itemsPerPage = 2;
    let selectedPageIndex = 0;

    let slice = this.paginate(itemsPerPage, selectedPageIndex);
    this.setState({ ...this.state, listColumns: columns, listItems: slice, selectedPageIndex: selectedPageIndex, itemsPerPage: itemsPerPage })

  }

  private getIdKey() {
    let idKey = "";
    if (this.state.query.endsWith("ies")) {
      idKey = `${this.state.query.substring(0, this.state.query.length - 3)}yid`
    } else {
      idKey = `${this.state.query.substring(0, this.state.query.length - 1)}id`
    }
    return idKey;
  }

  private loadColumns(idKey: any, responseJson: any): any[] {
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
        onColumnClick: this.onColumnClick
      })
    };
    addColumn(idKey);

    // Get only keys with name to filter columns on view and allow dynamic object types.
    Object.keys(responseJson.value[0]).filter(item => { return item.includes('name') && !item.includes('yomi') && !item.includes('address') }).forEach((column: any) => {
      addColumn(column);
    });

    return columns;
  }

  private loadItems(idKey: any, columns: any[], responseJson: any) {
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
    return mappedResponse;
  }

  private onChangeTextQuery = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: string | undefined): void => {
    this.setState({ ...this.state, query: text as string })
  };

  private onChangeTextFilter = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: string | undefined): void => {
    let filteredSet = this._allItems.filter(i => JSON.stringify(i).toLowerCase().indexOf((text as string).toLowerCase()) > -1)
    this.setState({
      listItems: text ? filteredSet : this._allItems,
    });
  };

  private onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
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
    const newItems = this.copyAndSort(listCollection, currColumn.fieldName!, currColumn.isSortedDescending);
    this.setState({
      listColumns: newColumns,
      listItems: newItems,
    });
  }

  private copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
    const key = columnKey as keyof T;
    return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
  }

  private getKey(item: any, index?: number): string {
    return item.key;
  }

  private paginate(page_size: number, page_number: number) {
    debugger;
    let slice = this._allItems.slice((page_number) * page_size, (page_number + 1) * page_size);
    return slice;
  }

  private onPageChange(ev: any) {
    debugger;
    let slice = this.paginate(this.state.itemsPerPage, ev);
    this.setState({ ...this.state, listItems: slice, selectedPageIndex: ev });
  }

  render() {
    const { query, listItems, listColumns, selectedPageIndex, itemsPerPage } = this.state;
    return (
      <Fabric>
        <Stack horizontal tokens={sectionStackTokens}>
          <Stack.Item grow>
            <TextField value={query} label="Api Query" onChange={this.onChangeTextQuery} />
          </Stack.Item>
          <Stack.Item shrink>
            <DefaultButton text="Submit Request" onClick={() => this.apiRequest(query)} allowDisabledFocus styles={buttonStyles} />
          </Stack.Item>
          <Stack.Item grow>
            <TextField label="Filter by any:" onChange={this.onChangeTextFilter} />
          </Stack.Item>
        </Stack>
        {
          listItems.length > 0 && listColumns.length > 0 ?
            <Fabric>
              <DetailsList
                items={listItems}
                columns={listColumns}
                getKey={this.getKey}
                setKey="multiple"
                layoutMode={DetailsListLayoutMode.justified}
              />
              <Pagination
                pageCount={this._allItems.length / (itemsPerPage as number)}
                itemsPerPage={(itemsPerPage as number)}
                totalItemCount={this._allItems.length}
                onPageChange={this.onPageChange.bind(this)}
                selectedPageIndex={selectedPageIndex}
                format={'buttons'}
                previousPageAriaLabel={'previous page'}
                nextPageAriaLabel={'next page'}
                firstPageAriaLabel={'first page'}
                lastPageAriaLabel={'last page'}
                pageAriaLabel={'page'}
                selectedAriaLabel={'selected'}
              />
            </Fabric>
            : null
        }
      </Fabric >
    )
  }
}


