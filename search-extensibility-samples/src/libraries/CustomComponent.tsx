import * as React from 'react';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import * as ReactDOM from 'react-dom';
import { DefaultButton, Dialog, ChoiceGroup, DialogFooter, PrimaryButton, DialogType, IChoiceGroupOption, Link, Icon, Label, TextField } from 'office-ui-fabric-react';

export interface IObjectParam {
    myProperty: string;
}

export interface ICustomComponentProps {

    productName?: string;
    productNumber?: string;

    /**
     * A sample object param
     */
    myObjectParam?: IObjectParam;

    /**
     * A sample date param
     */
    myDateParam?: Date;

    /**
     * A sample number param
     */
    myNumberParam?: number;

    /**
     * A sample boolean param
     */
    myBooleanParam?: boolean;
}

export interface ICustomComponenState {

    isDialogOpen: boolean;
}

export class CustomComponent extends React.Component<ICustomComponentProps, ICustomComponenState> {

    constructor(props) {
        super(props);

        this.state = {
            isDialogOpen: false
        };

        this.toogleDialog = this.toogleDialog.bind(this);
    }
    
    public render() {

        // Parse custom object
        const myObject: IObjectParam = this.props.myObjectParam;

        const modelProps = {
            isBlocking: false,
            styles: { main: { maxWidth: 450 } },
        };

        const dialogContentProps = {
        type: DialogType.largeHeader,
        title: `What\' wrong with the '${this.props.productName}' product?`,
        subText: `Please enter a quick description of the error for product ${this.props.productNumber}`,
        };

        return  <>
                    <div style={{display: 'flex', alignItems: 'center'}}>
                        <Icon iconName="WorkItemBug" />
                        <Link onClick={this.toogleDialog}>Report an issue</Link>    
                    </div>                         
                    <Dialog
                        hidden={!this.state.isDialogOpen}
                        onDismiss={this.toogleDialog}
                        dialogContentProps={dialogContentProps}
                        modalProps={modelProps}
                    >
                        <TextField multiline rows={3} />
                        <DialogFooter>
                            <PrimaryButton onClick={() => {
                                alert("Thank you!");
                            }} text="Submit" />
                            <DefaultButton onClick={this.toogleDialog} text="Cancel" />
                        </DialogFooter>
                    </Dialog>
                </>
               
    }

    public toogleDialog() {
        this.setState(
            {
                isDialogOpen: !this.state.isDialogOpen
            }
        )
    }
}

export class MyCustomComponentWebComponent extends BaseWebComponent {
   
    public constructor() {
        super(); 
    }
 
    public async connectedCallback() {
 
       let props = this.resolveAttributes();
       const customComponent = <CustomComponent {...props}/>;
       ReactDOM.render(customComponent, this);
    }    
}