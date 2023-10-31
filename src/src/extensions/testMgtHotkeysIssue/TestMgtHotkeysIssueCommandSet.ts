import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  type Command,
  type IListViewCommandSetExecuteEventParameters,
  type ListViewStateChangedEventArgs
} from '@microsoft/sp-listview-extensibility';

// Import MGT components
import { Providers, SharePointProvider } from '@microsoft/mgt-spfx';

// Import React components
import TestDialog from './components/testDialog/TestDialog';

export interface ITestMgtHotkeysIssueCommandSetProperties {
}

const LOG_SOURCE: string = 'TestMgtHotkeysIssueCommandSet';

export default class TestMgtHotkeysIssueCommandSet extends BaseListViewCommandSet<ITestMgtHotkeysIssueCommandSetProperties> {

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized TestMgtHotkeysIssueCommandSet');
    console.log('Initialized TestMgtHotkeysIssueCommandSet');

    // Initialize the MGT components infrastructure
    if (!Providers.globalProvider) {
      Providers.globalProvider = new SharePointProvider(this.context);
    }
  
    // initial state of the command's visibility
    const mgtTestCommand: Command = this.tryGetCommand('MGT_TEST_COMMAND');
    mgtTestCommand.visible = false;

    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);

    return Promise.resolve();
  }

  public async onExecute(event: IListViewCommandSetExecuteEventParameters): Promise<void> {
    switch (event.itemId) {
      case 'MGT_TEST_COMMAND':
        {
          // Prepare the settings to show the test dialog 
          const mgtTestDialog: TestDialog = new TestDialog();
          mgtTestDialog.cancel = async (): Promise<void> => { 
            alert('Cancel');
            console.log('Cancel');
            return; 
          };
          mgtTestDialog.save = async (people: string[]): Promise<void> => {
            alert('People selected!');
            console.log(people);
          };

          await mgtTestDialog.show();
        }          
        break;
      default:
        throw new Error('Unknown command');
    }
  }

  private _onListViewStateChanged = (args: ListViewStateChangedEventArgs): void => {
    Log.info(LOG_SOURCE, 'List view state changed');
    console.log('List view state changed');

    const mgtTestCommand: Command = this.tryGetCommand('MGT_TEST_COMMAND');
    if (mgtTestCommand) {
      // This command should be hidden unless exactly one row is selected.
      mgtTestCommand.visible = this.context.listView.selectedRows?.length === 1;
    }

    // TODO: Add your logic here

    // You should call this.raiseOnChage() to update the command bar
    this.raiseOnChange();
  }
}
