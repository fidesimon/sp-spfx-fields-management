import { FieldTypeKindEnum } from "../FieldTypeKindEnum";

export default class TextField{

    constructor() {
        this["@odata.type"] = "#SP.Field";
        this.FieldTypeKind = FieldTypeKindEnum.Text;
    }
    public '@odata.type'?: string;
    public DefaultValue: string;
    public Description: string;
    public EnforceUniqueValues: boolean;
    public FieldTypeKind: number;
    public Group: string;
    public InternalName: string;
    public MaxLength: number;
    public Required: boolean;
    public StaticName: string;
    public Title: string;
}