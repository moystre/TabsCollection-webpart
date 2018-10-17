import { Dropdown, IDropdownOption } from 'office-ui-fabric-react';
import * as React from 'react';
import { IMultiSelectPropInternal } from './PropertyPaneMultiSelect';
export interface IMultiSelectHostProp extends IMultiSelectPropInternal {
}
export interface IMultiSelectHostState {
    items?: any;
    selectedItems: any;
}
 
export class MultiSelectHost extends React.Component<IMultiSelectHostProp, IMultiSelectHostState>{    
    constructor(prop: IMultiSelectHostProp) {
        super(prop);
        this.state = ({
            items: [],
            selectedItems: this.props.selectedItemIds
        });
        this.onChangeMultiSelect = this.onChangeMultiSelect.bind(this); 
        this.props.onload()
            .then((items) => {
                this.setState({
                    items: items
                });
                // this._applyMultiSelect(this.props.selectedKey, this.props.selectedItemIds);
            });
    }

    public onChangeMultiSelect(item: IDropdownOption):void {
        let updatedSelectedItem = this.state.selectedItems ? this.copyArray(this.state.selectedItems) : [];
        if (item.selected) {
          // add the option if it's checked
          updatedSelectedItem.push(item.key);
        } else {
          // remove the option if it's unchecked
          const currIndex = updatedSelectedItem.indexOf(item.key);
          if (currIndex > -1) {
            updatedSelectedItem.splice(currIndex, 1);
          }
        }
        this.setState({
          selectedItems: updatedSelectedItem
        }, () => {
            let oldValues = this.props.properties[this.props.targetProperty];
            this.props.properties[this.props.targetProperty] = this.state.selectedItems;
            this.props.onPropChange(this.props.targetProperty, oldValues, this.state.selectedItems);
        });
    }

    public copyArray = (array: any[]): any[] => {
        const newArray: any[] = [];
        for (let i = 0; i < array.length; i++) {
          newArray[i] = array[i];
        }
        return newArray;
    }

    public render(): JSX.Element {        
        return (
            <div>
                <Dropdown
                label= {this.props.label}
                selectedKeys={ this.state.selectedItems }
                multiSelect
                options={ this.state.items }
                onChanged={ this.onChangeMultiSelect }
              />
            </div>
        );
    }
}