import * as React from 'react';
import styles from './WizardHost.module.scss';
import { IWizardHostProps } from './IWizardHostProps';
import { BaseWizard, WizardStep, IWizardStepValidationResult } from "../../../common/components/Wizard";
import { Button } from 'office-ui-fabric-react/lib/Button';
import { TextField } from "office-ui-fabric-react/lib/TextField";

// export enum MyWizardSteps {
//   None = 0,
//   FirstStep = 1,
//   SecondStep = 2,
//   ThirdStep = 4,
//   LastStep = 8
// }

export enum MyWizardSteps {
  None        = 0b0000,
  FirstStep   = 0b0001,
  SecondStep  = 0b0010,
  ThirdStep   = 0b0100,
  LastStep    = 0b1000
}


export class MyWizard extends BaseWizard<MyWizardSteps> {

}

export interface IWizardHostState {
  isWizardOpened: boolean;
  statusMessage: string;
  statusType: "OK" | "KO" | null;
  firstStepInput: string;
  thirdStepInput: string;
  wizardValidatingMessage: string;
}


export default class WizardHost extends React.Component<IWizardHostProps, IWizardHostState> {

  constructor(props: IWizardHostProps) {
    super(props);

    this.state = {
      isWizardOpened: false,
      statusMessage: null,
      statusType: null,
      firstStepInput: null,
      thirdStepInput: null,
      wizardValidatingMessage: 'Validating...'
    };
  }

  private _renderMyWizard() {

    return <MyWizard
      mainCaption="My Wizard"
      onCancel={() => this._closeWizard(false)}
      onCompleted={() => this._closeWizard(true)}
      onValidateStep={(step) => this._onValidateStep(step)}
      validatingMessage={this.state.wizardValidatingMessage}
    >
      <WizardStep caption="My first step" step={MyWizardSteps.FirstStep}>
        <div className={styles.wizardStep}>
          <h1>Hello from first step</h1>
          <TextField
            value={this.state.firstStepInput}
            placeholder="Type 'first' to validate the step"
            onChanged={(v) => this.setState({ firstStepInput: v })}></TextField>
        </div>
      </WizardStep>

      <WizardStep caption="My second step" step={MyWizardSteps.SecondStep}>
        <div className={styles.wizardStep}>
          <h1>Hello from second step</h1>
        </div>
      </WizardStep>

      <WizardStep caption="My third step" step={MyWizardSteps.ThirdStep}>
        <div className={styles.wizardStep}>
          <h1>Hello from third step</h1>
          <TextField
            value={this.state.thirdStepInput}
            placeholder="Type 'third' to validate the step (async validation)"
            onChanged={(v) => this.setState({ thirdStepInput: v })}></TextField>
        </div>
      </WizardStep>

      <WizardStep caption="My final step" step={MyWizardSteps.LastStep}>
        <div className={styles.wizardStep}>
          <h1>Hello from final step</h1>
        </div>
      </WizardStep>

      <div>
        Invalid element here, will be ignored
      </div>
    </MyWizard >;
  }

  private _openWizard() {
    this.setState({
      isWizardOpened: true
    });
  }

  private _closeWizard(completed: boolean = false) {
    this.setState({
      isWizardOpened: false,
      statusMessage: completed ? "The wizard has been completed" : "The wizard has been canceled",
      statusType: completed ? "OK" : "KO"
    });

    setTimeout(() => {
      this.setState({
        statusMessage: null,
        statusType: null
      });
    }, 3000);
  }

  private _onValidateStep(step: MyWizardSteps): IWizardStepValidationResult | Promise<IWizardStepValidationResult> {

    console.log('Validating step: ', step);
    let isValid = true;
    switch (step) {
      case MyWizardSteps.FirstStep:
        isValid = this.state.firstStepInput == 'first';
        return {
          isValidStep: isValid,
          errorMessage: !isValid ? "Your input to first step is invalid" : null
        };
      case MyWizardSteps.ThirdStep:

        return new Promise((resolve) => {
          isValid = this.state.thirdStepInput == 'third';
          setTimeout(() => {
            resolve({
              isValidStep: isValid,
              errorMessage: !isValid ? "Your input to third step is invalid" : null
            });
          }, 3000);
        });
      case MyWizardSteps.LastStep:
        this.setState({
          wizardValidatingMessage: 'Validating all the information you entered...'
        });
        return new Promise((resolve) => {
          isValid = this.state.thirdStepInput == 'third';
          setTimeout(() => {
            resolve({
              isValidStep: isValid,
              errorMessage: !isValid ? "One of your input is invalid" : null
            });
            this.setState({
              wizardValidatingMessage: null
            });
          }, 3000);
        });
      default:
        return { isValidStep: true };
    }
  }

  private _getStatusMessageCssClass(): string {
    switch (this.state.statusType) {
      case "OK":
        return `${styles.OK} ${styles.statusMessage}`;
      case "KO":
        return `${styles.KO} ${styles.statusMessage}`;
      default:
        return "";
    }
  }

  public render(): React.ReactElement<IWizardHostProps> {
    return (
      <div className={styles.wizardHost} >
        <div className={''}>
          <div className={styles.row}>
            <div className={styles.column}>
              {this.state.statusMessage && <div className={this._getStatusMessageCssClass()}>
                {this.state.statusMessage}
              </div>}
              {this.state.isWizardOpened
                ? this._renderMyWizard()
                : <div>
                  <Button text="Open wizard" onClick={() => this._openWizard()} />
                </div>}
            </div>
          </div>
        </div>
      </div>
    );
  }
}
