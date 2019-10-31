import * as React from "react";
import { ActionButton } from "office-ui-fabric-react/lib/Button";
import { Pivot, PivotItem, IPivotItemProps } from "office-ui-fabric-react/lib/Pivot";
import styles from "./Wizard.module.scss";

export interface IWizardStepProps<TStep extends number> extends IPivotItemProps {
    step: TStep;
    caption: string;
}

export class WizardStep<TStep extends number> extends React.Component<IWizardStepProps<TStep>, {}> {

}

export interface IWizardStepValidationResult {
    isValidStep: boolean;
    errorMessage?: string;
}

export interface IWizardProps<TStep extends number> {
    defaultCurrentStep?: TStep;
    onValidateStep?: (currentStep: TStep) => IWizardStepValidationResult | Promise<IWizardStepValidationResult>;
    onCompleted?: () => void;
    onCancel?: () => void;
    nextButtonLabel?: string;
    previousButtonLabel?: string;
    cancelButtonLabel?: string;
    finishButtonLabel?: string;
    validatingMessage?: string;
    mainCaption?: string;
}

export interface IWizardState<TStep extends number> {
    currentStep: TStep;
    completedSteps: TStep;
    errorMessage: string;
    isValidatingStep: boolean;
}

const DEFAULT_NEXT_BUTTON_LABEL = "Next";
const DEFAULT_PREVIOUS_BUTTON_LABEL = "Previous";
const DEFAULT_FINISH_BUTTON_LABEL = "Finish";
const DEFAULT_CANCEL_BUTTON_LABEL = "Cancel";
const DEFAULT_VALIDATING_MESSAGE = "Validating step...";

export abstract class BaseWizard<TStep extends number> extends React.Component<IWizardProps<TStep>, IWizardState<TStep>> {

    constructor(props: IWizardProps<TStep>) {
        super(props);

        this.state = {
            currentStep: props.defaultCurrentStep || this.firstStep,
            completedSteps: null,
            errorMessage: null,
            isValidatingStep: false
        };
    }

    private renderSteps() {
        const stepChildren = React.Children.toArray(this.props.children)
            .filter((reactChild: React.ReactElement) => reactChild.type && (reactChild.type as any).name == 'WizardStep' && reactChild.props.step);

        if (stepChildren.length == 0) {
            throw new Error("The specified wizard steps are not valid");
        }

        return stepChildren
            .map((reactChild: React.ReactElement) => {
                return <PivotItem key={`WizardStep__${reactChild.props.step}`}
                    itemKey={reactChild.props.step.toString()}
                    headerText={reactChild.props.caption}
                    headerButtonProps={{}} >
                    {reactChild.props.children}
                </PivotItem>;
            });
    }

    private get firstStep(): TStep {
        const stepValues = React.Children.toArray(this.props.children)
            .filter((c: React.ReactElement) => c.props.step as number > 0)
            .map((c: React.ReactElement) => c.props.step as number);
        if (stepValues.length < 1) {
            throw new Error("The specified step values are invalid. First step value must be higher than 0");
        }
        return Math.min(...stepValues) as TStep;
    }

    private get lastStep(): TStep {
        const stepValues = React.Children.toArray(this.props.children)
            .filter((c: React.ReactElement) => c.props.step as number > 0)
            .map((c: React.ReactElement) => c.props.step as number);
        if (stepValues.length < 1) {
            throw new Error("The specified step values are invalid. First step value must be higher than 0");
        }
        return Math.max(...stepValues) as TStep;
    }

    private _validateWithCallback = (validationCallback: (validationResult: IWizardStepValidationResult) => void) => {

        if (!validationCallback) {
            return;
        }

        const validationResult = this._validateStep(this.state.currentStep);
        if (typeof (validationResult as Promise<IWizardStepValidationResult>).then === "function") {
            this.setState({
                isValidatingStep: true,
                errorMessage: null
            });
            const promiseResult = validationResult as Promise<IWizardStepValidationResult>;
            promiseResult.then(result => {
                validationCallback(result);
            }).catch(error => {
                if (error as string) {
                    validationCallback({
                        isValidStep: false,
                        errorMessage: error
                    });
                }
            });
        }
        else {
            const directResult = validationResult as IWizardStepValidationResult;
            if (!directResult) {
                throw new Error("The validation result has unexpected format.");
            }
            validationCallback(directResult);
        }
    }

    private _goToStep = (step: TStep, completedSteps?: TStep, skipValidation: boolean = false) => {

        if (!skipValidation) {
            this._validateWithCallback(result => {
                if (result.isValidStep) {

                    this.setState({
                        currentStep: step,
                        completedSteps,
                        errorMessage: null,
                        isValidatingStep: false
                    });
                } else {
                    this.setState({
                        errorMessage: result.errorMessage,
                        isValidatingStep: false
                    });
                }
            });
        } else {
            this.setState({ currentStep: step, completedSteps });
        }
    }

    private _validateStep = (step: TStep) => {
        if (this.props.onValidateStep) {
            return this.props.onValidateStep(step);
        }

        return {
            isValidStep: true,
            errorMessage: null
        };
    }

    private get hasNextStep(): boolean {
        return this.state.currentStep < this.lastStep;
    }

    private get hasPreviousStep(): boolean {
        return this.state.currentStep > this.firstStep;
    }

    private _goToNextStep = () => {
        let completedWizardSteps = (this.state.completedSteps | this.state.currentStep) as TStep;
        const nextStep = ((this.state.currentStep as number) << 1) as TStep;
        console.log("Current step: ", this.state.currentStep, " next step: ", nextStep);
        this._goToStep(nextStep, completedWizardSteps);
    }

    private _goToPreviousStep = () => {
        const previousStep = ((this.state.currentStep as number) >> 1) as TStep;
        console.log("Current step: ", this.state.currentStep, " previous step: ", previousStep);
        this._goToStep(previousStep, null, true);
    }


    private _cancel = () => {
        if (this.props.onCancel) {
            this.props.onCancel();
        }
    }

    private _finish = () => {
        this._validateWithCallback((result) => {
            if (result.isValidStep) {
                if (this.props.onCompleted) {
                    this.props.onCompleted();
                }
            } else {
                this.setState({
                    errorMessage: result.errorMessage,
                    isValidatingStep: false
                });
            }
        });
    }

    private get cancelButton(): JSX.Element {
        return <ActionButton iconProps={{ iconName: "Cancel" }} text={this.props.cancelButtonLabel || DEFAULT_CANCEL_BUTTON_LABEL} onClick={this._cancel} />;
    }

    private get previousButton(): JSX.Element {
        if (this.hasPreviousStep) {
            return <ActionButton iconProps={{ iconName: "ChevronLeft" }} text={this.props.previousButtonLabel || DEFAULT_PREVIOUS_BUTTON_LABEL} onClick={this._goToPreviousStep} />;
        }

        return null;
    }

    private get nextButton(): JSX.Element {
        if (this.hasNextStep) {
            return <ActionButton iconProps={{ iconName: "ChevronRight" }} text={this.props.nextButtonLabel || DEFAULT_NEXT_BUTTON_LABEL} onClick={this._goToNextStep} />;
        }

        return null;
    }

    private get finishButton(): JSX.Element {
        if (!this.hasNextStep) {
            return <ActionButton iconProps={{ iconName: "Save" }} text={this.props.finishButtonLabel || DEFAULT_FINISH_BUTTON_LABEL} onClick={this._finish} />;
        }

        return null;
    }

    public render(): React.ReactElement<IWizardProps<TStep>> {
        return <div className={styles.wizardComponent}>
            {this.props.mainCaption && <h1>{this.props.mainCaption}</h1>}
            <Pivot selectedKey={this.state.currentStep.toString()}>
                {this.renderSteps()}
            </Pivot>
            {this.state.isValidatingStep && <div>{this.props.validatingMessage || DEFAULT_VALIDATING_MESSAGE}</div>}
            {this.state.errorMessage && <div className={styles.error}>{this.state.errorMessage}</div>}

            <div className={styles.row}>
                <div className={`${styles.halfColumn} ${styles.lefted}`}>
                    {this.cancelButton}
                </div>
                <div className={`${styles.halfColumn} ${styles.righted}`}>
                    {this.previousButton}
                    {this.nextButton}
                    {this.finishButton}
                </div>
            </div>


        </div>;
    }
}