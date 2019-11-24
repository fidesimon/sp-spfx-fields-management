import { FieldTypeKindEnum } from "../FieldTypeKindEnum";

export default class TextField{

    constructor() {
        this["@odata.type"] = "#SP.Field";
        this.FieldTypeKind = FieldTypeKindEnum.Text;
    }
    '@odata.type'?: string;
    DefaultValue: string;
    Description: string;
    EnforceUniqueValues: boolean;
    FieldTypeKind: number;
    Group: string;
    InternalName: string;
    MaxLength: number;
    Required: boolean;
    StaticName: string;
    Title: string;
}