import * as React from 'react';
import styles from './HelloWorld.module.scss';
import { IHelloWorldProps } from './IHelloWorldProps';
import { escape } from '@microsoft/sp-lodash-subset';
import CsvParse from '@vtex/react-csv-parse';
import * as $ from 'jquery';
import { sp, ItemAddResult } from "@pnp/sp";

export default class HelloWorld extends React.Component<IHelloWorldProps, {}> {
  
  public constructor(props: IHelloWorldProps) {
    super(props);
    this.state = {
      data: null,
      error: null,
    };
  }

  public handleData = data => {
    this.setState({ data });

    $.each(data, (e, val) => {

      // we can also use filter as a supported odata operation, but this will likely fail on large lists
      sp.web.lists.getByTitle("test").items.filter("Title eq '" + val.RecordID + "'").getAll().then((allItems: any[]) => {
        // how many did we get
        if (allItems.length > 0) {
          $.each(allItems, (e1, val1) => {
            console.log("Update");
            let list = sp.web.lists.getByTitle("test");
            list.items.getById(allItems[e1].Id).update({
              Title: val.RecordID,
              DefectID: val.DefectID,
              Description: val.Description,
              FullDescription: val.FullDescription,
              Severity: val.Severity,
              Status: val.Status,
              AssignedTo: val.AssignedTo,
              Tester: val.Tester,
              TestPhase: val.TestPhase
            }).then(i => {
              console.log(i);
            });
          });
        }
        else {
          // add an item to the list
          sp.web.lists.getByTitle("test").items.add({
            Title: val.RecordID,
            DefectID: val.DefectID,
            Description: val.Description,
            FullDescription: val.FullDescription,
            Severity: val.Severity,
            Status: val.Status,
            AssignedTo: val.AssignedTo,
            Tester: val.Tester,
            TestPhase: val.TestPhase
          }).then((iar: ItemAddResult) => {
            console.log(iar);
          });
        }
      });
    });
  }


  public handleError = error => {
    this.setState({ error });
  }

  public render(): React.ReactElement<IHelloWorldProps> {
    const keys = [
      'RecordID',
      'DefectID',
      'Description', 'FullDescription', 'Severity', 'Status', 'AssignedTo', 'Tester', 'TestPhase'
    ];
    return (
      <CsvParse
        keys={keys}
        onDataUploaded={this.handleData}
        onError={this.handleError}
        render={onChange => <input type="file" onChange={onChange} />}
      />
    );
  }
}
