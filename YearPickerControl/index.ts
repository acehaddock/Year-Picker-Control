import {IInputs, IOutputs} from "./generated/ManifestTypes";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { DatePicker, IDatePickerStrings, mergeStyleSets, DayOfWeek } from 'office-ui-fabric-react';

export class YearPickerControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    private notifyOutputChanged: () => void;
    private container: HTMLDivElement;
    private context: ComponentFramework.Context<IInputs>;

    /**
     * Empty constructor.
     */
    constructor()
    {

    }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement): void
    {
        // Add control initialization code
        this.notifyOutputChanged = notifyOutputChanged;
        this.container = container;
        this.context = context;
    }
    /**
     * Method to render the control
     */
    public render(): void {
        ReactDOM.render(
            React.createElement(YearPicker, {
                selectedYear: this.context.parameters.dateInput.raw,
                onSelectYear: this.onSelectYear.bind(this)
            }),
            this.container
        );
    }

    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void
    {
        // Add code to update control view
        this.render();
    }
    
    /**
     * Method to handle year selection
     * @param year The selected year
     */
        private onSelectYear(year: number): void {
            this.context.parameters.dateInput.raw = year.toString();
            this.notifyOutputChanged();
        }
    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
     */
    public getOutputs(): IOutputs
    {
        return {};
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void
    {
        // Add code to cleanup control if necessary
        ReactDOM.unmountComponentAtNode(this.container);
    }
}

/**
 * Interface for YearPicker properties
 */
interface IYearPickerProps {
    selectedYear: string;
    onSelectYear: (year: number) => void;
}

/**
 * YearPicker component using Fluent UI DatePicker
 */
const YearPicker: React.FunctionComponent<IYearPickerProps> = (props) => {

    // Convert selected year to a Date object
    const selectedDate: Date = new Date(parseInt(props.selectedYear));

    // Handler for date selection change
    const onSelectDate = (date: Date | null | undefined) => {
        if (date) {
            const selectedYear = date.getFullYear();
            props.onSelectYear(selectedYear);
        }
    };

    // Customize DatePicker strings to only show year
    const datePickerStrings: IDatePickerStrings = {
        months: [],
        shortMonths: [],
        days: [],
        shortDays: [],
        goToToday: "",
        prevMonthAriaLabel: "",
        nextMonthAriaLabel: "",
        prevYearAriaLabel: "",
        nextYearAriaLabel: "",
        isRequiredErrorMessage: "",
        invalidInputErrorMessage: "",
    };

    // Render the YearPicker control
    return (
        <div>
            <DatePicker
                label=""
                value={selectedDate}
                formatDate={date => `${date.getFullYear()}`}
                strings={datePickerStrings}
                allowTextInput={false}
                showMonthPickerAsOverlay={false}
                showWeekNumbers={false}
                firstDayOfWeek={DayOfWeek.Sunday}
                onSelectDate={onSelectDate}
            />
        </div>
    );
};