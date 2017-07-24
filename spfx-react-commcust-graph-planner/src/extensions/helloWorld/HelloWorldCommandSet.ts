import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  IListViewCommandSetRefreshEventParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';

import { Dialog } from '@microsoft/sp-dialog';
import PlannerDialog from './PlannerDialog';

import { GraphHttpClient, GraphClientResponse, IGraphHttpClientOptions } from '@microsoft/sp-http';

export interface IHelloWorldCommandSetProperties {
  // planId in Planner
  planId: string;
}

const LOG_SOURCE: string = 'HelloWorldCommandSet';

export default class HelloWorldCommandSet
  extends BaseListViewCommandSet<IHelloWorldCommandSetProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized HelloWorldCommandSet');
    return Promise.resolve<void>();
  }

  @override
  public onRefreshCommand(event: IListViewCommandSetRefreshEventParameters): void {
    // show button when one item is selected
    event.visible = event.selectedRows.length === 1;
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {

    switch (event.commandId) {
      // if command 'Add To O365 Planner' is clicked
      case 'CMDAddToPlanner':
        // create PlannerDialog and put Title field to it
        const dialog: PlannerDialog = new PlannerDialog();
        dialog.title = event.selectedRows[0].getValueByName('Title').toString();

        // show dialog
        dialog.show().then(() => {
          // if Title is not empty
          if (dialog.title != "") {
            // get bucketID for specific PlanID from Planner (via MS Graph)
            this.context.graphHttpClient.get("beta/plans/" + this.properties.planId + "/buckets?$select=id", GraphHttpClient.configurations.v1)
            .then((response: GraphClientResponse): Promise<any> => {
              return response.json();
            })
            .then((bucketData: any): void => {
              if (bucketData.error) {
                Dialog.alert(bucketData.error.message);
              }
              else {
                // Currently via GraphHttpClient you cannot access to 'v1.0/me' for any additional information about current user (for example ID which you need it later).
                // For that reason my userID is hardcoded below. You can use predefined dropdown with userIDs and Display Names or ADAL JS with implicit OAuth flow instead.

                //this.context.graphHttpClient.get("v1.0/me?$select=id", GraphHttpClient.configurations.v1)
                //.then((response: GraphClientResponse): Promise<any> => {
                //  return response.json();
                //})
                //.then((meData: any): void => {
                //  if (meData.error) {
                //    Dialog.alert(meData.error.message);
                //  }
                //  else {
                //    var myId = meData.value[0].id;

                    var options : IGraphHttpClientOptions = {
                      method: "POST",
                      body: JSON.stringify({ 
                        planId: this.properties.planId, 
                        bucketId: bucketData.value[0].id,
                        title: dialog.title,
                        assignments: {
                          "[your-userID-hardcoded]": { // myId: {
                            "@odata.type": "#microsoft.graph.plannerAssignment",
                            "orderHint": " !"
                          }
                        }
                      })
                    };

                    // add task to planner
                    this.context.graphHttpClient.fetch("v1.0/planner/tasks", GraphHttpClient.configurations.v1, options)
                    .then((response: GraphClientResponse): Promise<any> => {
                      return response.json();
                    })
                    .then((taskData: any): void => {
                      if (taskData.error) {
                        Dialog.alert(taskData.error.message);
                      }
                      else {
                        Dialog.alert("Task successfully added to O365 Planner! :)");
                      }
                    });

                //  }
                //});
              }
            });
          }
        });
        break;
      default:
        throw new Error('Unknown command');
    }
  }
}
