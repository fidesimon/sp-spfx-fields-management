import * as React from 'react';
import { DetailsList, DetailsListLayoutMode, Selection, SelectionMode, DetailsColumnBase } from 'office-ui-fabric-react/lib/DetailsList';
import { IColumn } from 'office-ui-fabric-react/lib/components/DetailsList/DetailsList.types';
import { mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';
import { IGroup } from '../Group';
import { ISPField } from '../SPField';
import { Icon, CommandBar } from 'office-ui-fabric-react';


interface IDisplayFieldsProps {
    fields: IGroup;
    removeFieldHandler: Function;
    addFieldHandler: Function;
}

const classNames = mergeStyleSets({
    headerIcon: {
        padding: 0,
        fontSize: '16px'
    },
    iconCell: {
        textAlign: 'center',
        selectors: {
            '&:before': {
                content: '.',
                display: 'inline-block',
                verticalAlign: 'middle',
                height: '100%',
                width: '0px',
                visibility: 'hidden'
            }
        }
    }
});
const controlStyles = {
    root: {
        margin: '0 30px 20px 0',
        maxWidth: '300px'
    }
};

export interface IDetailsListDocumentsExampleState {
    columns: IColumn[];
    items: IDocument[];
    selectionDetails: string;
    isModalSelection: boolean;
    isCompactMode: boolean;
    announcedMessage?: string;
}

export interface IDocument {
    key: string;
    name: string;
    value: string;
    internalName: string;
    typeDisplayName: string;
    fieldId: string;
    removable: string;
    group: string;
}


export default class DisplayFields extends React.Component<IDisplayFieldsProps, IDetailsListDocumentsExampleState>{
    private _selection: Selection;
    private _allItems: IDocument[];
    constructor(props: IDisplayFieldsProps) {
        super(props);
        this._allItems = _generateDocuments(this.props.fields);

        const columns: IColumn[] = [
            {
                key: 'column2',
                name: 'Title',
                fieldName: 'name',
                minWidth: 90,
                isRowHeader: true,
                isResizable: true,
                isSorted: true,
                isSortedDescending: false,
                sortAscendingAriaLabel: 'Sorted A to Z',
                sortDescendingAriaLabel: 'Sorted Z to A',
                onColumnClick: this._onColumnClick,
                data: 'string',
                isPadded: true,
                onRender: (item: IDocument) => {
                    return <span key={item.key}>{item.name}</span>;
                }
            },
            {
                key: 'column3',
                name: 'Field Type',
                fieldName: 'typeDisplayName',
                minWidth: 90,
                isResizable: true,
                onColumnClick: this._onColumnClick,
                data: 'string',
                onRender: (item: IDocument) => {
                    return <span>{item.typeDisplayName}</span>;
                },
                isPadded: true
            },
            {
                key: 'column4',
                name: 'Internal Name',
                fieldName: 'internalName',
                minWidth: 90,
                isResizable: true,
                isCollapsible: true,
                data: 'string',
                onColumnClick: this._onColumnClick,
                onRender: (item: IDocument) => {
                    return <span>{item.internalName}</span>;
                },
                isPadded: true
            },
            {
                key: 'column5',
                name: 'ID',
                fieldName: 'fieldId',
                minWidth: 90,
                isResizable: true,
                isCollapsible: true,
                data: 'string',
                onColumnClick: this._onColumnClick,
                onRender: (item: IDocument) => {
                    return <span>{item.fieldId}</span>;
                }
            },
            {
                key: 'column6',
                name: 'Delete',
                className: classNames.iconCell,
                iconClassName: classNames.headerIcon,
                ariaLabel: 'Delete Item',
                iconName: 'Delete',
                isIconOnly: true,
                fieldName: 'removable',
                minWidth: 16,
                maxWidth: 16,
                onRender: (item: IDocument) => {
                    return <Icon iconName="Delete" style={{ color: item.removable == "T" ? "#ff0000" : "#d8d8d8" }} onClick={() => { this.deleteItem(item.fieldId, item.group) }} />
                }
            }
        ];

        this._selection = new Selection({
            onSelectionChanged: () => {
                this.setState({
                    selectionDetails: this._getSelectionDetails()
                });
            }
        });

        this.state = {
            items: this._allItems,
            columns: columns,
            selectionDetails: this._getSelectionDetails(),
            isModalSelection: false,
            isCompactMode: false,
            announcedMessage: undefined
        };
    }

    protected async deleteItem(fieldId: string, group: string) {
        if (await this.props.removeFieldHandler(fieldId, group)) {
            let items = this.state.items.filter(n => n.fieldId !== fieldId);
            this.setState({ items: items });
        } else {
            alert("Failed to delete the item");
        }
    }

    protected async addFieldHandler(groupName: string){
        let data = await this.props.addFieldHandler(groupName);
    }

    render() {
        const { columns, isCompactMode, items, selectionDetails, isModalSelection, announcedMessage } = this.state;

        return (<>
            <CommandBar
                items={
                    [{
                        key: 'newItem',
                        text: 'New',
                        cacheKey: 'myCacheKey', 
                        iconProps: { iconName: 'Add' },
                        onClick: () => { this.addFieldHandler(this.props.fields.Name) }
                    },
                    {
                        key: 'delete',
                        text: 'Delete',
                        iconProps: { iconName: 'Delete' },
                        disabled: true
                    }]
                }
                ariaLabel="Use left and right arrow keys to navigate between commands"
            />
            <DetailsList
                items={items}
                compact={false}
                columns={columns}
                selectionMode={SelectionMode.multiple}
                setKey="multiple"
                layoutMode={DetailsListLayoutMode.justified}
                isHeaderVisible={true}
                selectionPreservedOnEmptyClick={true}
                enterModalSelectionOnTouch={true}
                ariaLabelForSelectionColumn="Toggle selection"
                ariaLabelForSelectAllCheckbox="Toggle selection for all items"
                checkButtonAriaLabel="Row checkbox"
            />
        </>
        );
    }

    public componentDidUpdate(previousProps: any, previousState: IDetailsListDocumentsExampleState) {
        if (previousState.isModalSelection !== this.state.isModalSelection && !this.state.isModalSelection) {
            this._selection.setAllSelected(false);
        }
    }

    private _getKey(item: any, index?: number): string {
        return item.key;
    }

    private _onChangeCompactMode = (ev: React.MouseEvent<HTMLElement>, checked: boolean): void => {
        this.setState({ isCompactMode: checked });
    };

    private _onChangeModalSelection = (ev: React.MouseEvent<HTMLElement>, checked: boolean): void => {
        this.setState({ isModalSelection: checked });
    };

    private _onChangeText = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: string): void => {
        this.setState({
            items: text ? this._allItems.filter(i => i.name.toLowerCase().indexOf(text) > -1) : this._allItems
        });
    };

    private _onItemInvoked(item: any): void {
        alert(`Item invoked: ${item.name}`);
    }

    private _getSelectionDetails(): string {
        const selectionCount = this._selection.getSelectedCount();

        switch (selectionCount) {
            case 0:
                return 'No items selected';
            case 1:
                return '1 item selected: ' + (this._selection.getSelection()[0] as IDocument).name;
            default:
                return `${selectionCount} items selected`;
        }
    }

    private _onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
        const { columns, items } = this.state;
        const newColumns: IColumn[] = columns.slice();
        const currColumn: IColumn = newColumns.filter(currCol => column.key === currCol.key)[0];
        newColumns.forEach((newCol: IColumn) => {
            if (newCol === currColumn) {
                currColumn.isSortedDescending = !currColumn.isSortedDescending;
                currColumn.isSorted = true;
                this.setState({
                    announcedMessage: `${currColumn.name} is sorted ${currColumn.isSortedDescending ? 'descending' : 'ascending'}`
                });
            } else {
                newCol.isSorted = false;
                newCol.isSortedDescending = true;
            }
        });
        const newItems = _copyAndSort(items, currColumn.fieldName!, currColumn.isSortedDescending);
        this.setState({
            columns: newColumns,
            items: newItems
        });
    };
}

function _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
    const key = columnKey as keyof T;
    return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
}

function _generateDocuments(propItems: IGroup) {
    let itanz: ISPField[] = propItems.Fields;
    const items: IDocument[] = [];
    for (let i = 0; i < itanz.length; i++) {
        items.push({
            key: i.toString(),
            name: itanz[i].Title,
            value: itanz[i].Title,
            internalName: itanz[i].StaticName,
            typeDisplayName: itanz[i].TypeDisplayName,
            fieldId: itanz[i].Id,
            removable: itanz[i].CanBeDeleted ? (itanz[i].SchemaXml.indexOf("schemas.microsoft") === -1 ? "T" : "F") : "F",
            group: itanz[i].Group
        });
    }
    return items;
}