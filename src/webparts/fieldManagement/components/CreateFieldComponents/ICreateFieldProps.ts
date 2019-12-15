import { IDropdownOption } from "office-ui-fabric-react";

export interface ICreateFieldProps {
    fieldTypeOptions: IDropdownOption[];
    saveButtonHandler: Function;
    cancelButtonHandler: Function;
    groupName: string;
    onFieldTypeChange: Function;
}