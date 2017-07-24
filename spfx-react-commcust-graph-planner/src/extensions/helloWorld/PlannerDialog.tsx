import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import {
  autobind,
  TextField,
  PrimaryButton,
  Button,
  DialogFooter
} from 'office-ui-fabric-react';
import { DialogContent } from '@microsoft/sp-dialog';

interface IPlannerDialogContentProps {
  close: () => void;
  submit: (title: string) => void;
  title?: string;
}

class PlannerDialogContent extends React.Component<IPlannerDialogContentProps, {}> {
  private _title: string;

  constructor(props) {
    super(props);

    // get default title
    this._title = props.title;
  }

  public render(): JSX.Element {
    // UI
    return <DialogContent
      title='Planner Task Details'
      subText='Check details below:'
      onDismiss={this.props.close}
      showCloseButton={true}
    >
      
      <TextField label='Title' required={ true } multiline autoAdjustHeight value={ this._title } onChanged={ this._onChanged } />
      <DialogFooter>
        <Button text='Cancel' title='Cancel' onClick={this.props.close} />
        <PrimaryButton text='OK' title='OK' onClick={() => { this.props.submit(this._title); }} />
      </DialogFooter>
    </DialogContent>;
  }

  @autobind
  private _onChanged(text: string) {
    this._title = text;
  }
}

export default class PlannerDialog extends BaseDialog {
  public title: string;

  public render(): void {
    ReactDOM.render(<PlannerDialogContent
      close={ this._close }
      title={ this.title }
      submit={ this._submit }
    />, this.domElement);
  }

  public getConfig(): IDialogConfiguration {
    return {
      isBlocking: false
    };
  }

  // onClose event
  @autobind
  private _close(): void {
    this.title = "";

    this.close();
  }

  // onSubmit event
  @autobind
  private _submit(title: string): void {
    this.title = title;

    this.close();
  }
}
