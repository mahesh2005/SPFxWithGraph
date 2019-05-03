import * as React from 'react';
import styles from './AzureAdGroupViewer.module.scss';
import { IAzureAdGroupViewerProps } from './IAzureAdGroupViewerProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Group,User } from '@microsoft/microsoft-graph-types';
import { graph } from '@pnp/graph';
import { createRef } from 'office-ui-fabric-react/lib/Utilities';
import {
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  IDetailsList,
  SelectionMode
} from 'office-ui-fabric-react/lib/DetailsList';
import { sp } from "@pnp/sp";
import { EnvironmentType } from '@microsoft/sp-core-library';

export interface IState {
  groups: Group[], 
  grpMembers:User[]
}

export default class AzureAdGroupViewer extends React.Component<IAzureAdGroupViewerProps,IState> {
  private _detailsList = createRef<IDetailsList>();

  _columns: IColumn[] = [
    {
      key: 'Name',
      name: 'Name',
      fieldName: 'Title',
      minWidth: 50,
      maxWidth: 100,
      isResizable: true
    },
    {
      key: 'Description',
      name: 'Description',
      fieldName: 'LoginName',
      minWidth: 50,
      maxWidth: 150,
      isResizable: true
    },
    {
      key: 'button',
      name: '',
      fieldName: '',
      minWidth: 50,
      maxWidth: 150,
      isResizable: true,    
      onRender:(item)=> {
        return <button onClick={() => this.onClick(item, event)} >Get Members</button>;
      }
    }
    
  ];

  _columnsUsers: IColumn[] = [
    {
      key: 'Name',
      name: 'Name',
      fieldName: 'displayName',
      minWidth: 50,
      maxWidth: 100,
      isResizable: true
    }  ,
    {
      key: 'Email',
      name: 'Email',
      fieldName: 'mail',
      minWidth: 50,
      maxWidth: 100,
      isResizable: true
    } ,
    {
      key: 'JobTitle',
      name: 'Job Title',
      fieldName: 'jobTitle',
      minWidth: 50,
      maxWidth: 100,
      isResizable: true
    }  

  ];

  constructor(props: any) {
    super(props);

    this.state = {
      groups: [],   
      grpMembers:[]
    };
    this.onClick = this.onClick.bind(this);
  }

  onClick = (item,event) => {       
    let gid = item.LoginName;
    gid = gid.substr(gid.lastIndexOf("|") + 1, gid.length - gid.lastIndexOf("|"));
    graph.groups.getById(gid).members.get().then(members => {    
      this.setState({
        grpMembers:members
      })  
    }); 
    event.preventDefault();    
  }

  public componentDidMount(): void {
    sp.web.siteGroups.getByName("Communication site Members").users.get().then(users=>{
      this.setState({
        groups: users   
      });        
    });  
  }

  public render(): React.ReactElement<IAzureAdGroupViewerProps> {   
    var style1= {
      'padding-top':'50px'
    };
    return (
      <div className={ styles.azureAdGroupViewer }>
      <div style={style1}> Groups:</div>
        <div>
          <DetailsList
            componentRef={this._detailsList}
            items={this.state.groups}
            columns={this._columns}
            setKey="set"
            layoutMode={DetailsListLayoutMode.justified}
            selectionMode={SelectionMode.none}
          />
        </div>
        <div style={style1}> Group Members:</div>
        <div>
          <DetailsList
            componentRef={this._detailsList}
            items={this.state.grpMembers}
            columns={this._columnsUsers}
            setKey="set"
            layoutMode={DetailsListLayoutMode.justified}
            selectionPreservedOnEmptyClick={true}
            ariaLabelForSelectionColumn="Toggle selection"
            ariaLabelForSelectAllCheckbox="Toggle selection for all items"
          />
        </div>
      </div>
    );
  }
}
