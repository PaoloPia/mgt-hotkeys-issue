import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import styles from './TestDialog.module.scss';

import {
    DefaultButton,
    PrimaryButton,
    DialogFooter,
    DialogContent,
    Label
} from '@fluentui/react/lib';
import { PeoplePicker } from '@microsoft/mgt-react/dist/es6/spfx';
import { PersonType } from '@microsoft/mgt-spfx';

import { ITestDialogContentProps } from './ITestDialogContentProps';
import { ITestDialogContentState } from './ITestDialogContentState';

class TestDialogContent extends
  React.Component<ITestDialogContentProps, ITestDialogContentState> {

    /**
     *
     */
    public constructor(props: ITestDialogContentProps) {
        super(props);
        
        this.state = {
            people: []
        };
    }

    public render(): JSX.Element {

        const {
            cancel,
            save
        } = this.props;

        return (<div className={styles.testDialogRoot}>
            <DialogContent
                title="MGT Test Dialog"
                subText="Test dialog to test MGT People Picker controls in a Dialog"
                onDismiss={cancel}>

            <div className={styles.testDialogContent}>
                <div className="ms-Grid">
                    <div className={`ms-Grid-row ${styles.rowSpacing}`}>
                        <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                            <Label>Pick people</Label>
                            <PeoplePicker type={PersonType.person} placeholder="Write people names ..."
                                selectionChanged={this._onPeoplePickerSelectionChanged} />
                        </div>
                    </div>
                </div>
            </div>
            <DialogFooter>
                <DefaultButton text="Cancel"
                    title="Cancel" onClick={async () => await cancel()} />
                <PrimaryButton text="OK" 
                    title="OK" onClick={async () => await save(this.state.people)} />
            </DialogFooter>
        </DialogContent>
    </div>);
    }

   private _onPeoplePickerSelectionChanged = (e: Event): void => {
        const recipients: { id: string, userPrincipalName: string}[] = (e as CustomEvent).detail;
        this.setState({
            people: recipients.map((i) => {return i.userPrincipalName})
        });
    }
}

export default class TestDialog extends BaseDialog {

    public cancel: () => Promise<void>;
    public save: (people: string[]) => Promise<void>;

    /**
     * Constructor to initialize fundamental properties
     */
    public constructor() {
        super();        
    }

    public render(): void {
        ReactDOM.render(<TestDialogContent
            cancel={ this._cancel }
            save={ this._save }
        />, this.domElement);
    }

    public getConfig(): IDialogConfiguration {
        return {
            isBlocking: false
        };
    }

    private _cancel = async (): Promise<void> => {
        await this.close();
        await this.cancel();
        return;   
    }

    private _save = async (people: string[]): Promise<void> => {
        await this.close();
        await this.save(people);
        return;   
    }

    protected onAfterClose(): void {
        super.onAfterClose();

        // Clean up the element for the next dialog
        ReactDOM.unmountComponentAtNode(this.domElement);
    }
}