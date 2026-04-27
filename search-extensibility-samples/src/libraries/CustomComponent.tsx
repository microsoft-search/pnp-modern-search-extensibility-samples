import * as React from 'react';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import * as ReactDOM from 'react-dom';
import { PageContext } from '@microsoft/sp-page-context';
import { AadTokenProviderFactory, MSGraphClientFactory, SPHttpClient } from '@microsoft/sp-http';
import { BaseComponentContext } from '@microsoft/sp-component-base';

export interface IObjectParam {
    myProperty: string;
}

export interface ICustomComponentProps {

    /**
     * A sample string param
     */
    myStringParam?: string;

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

    /**
     * A BaseComponentContext object
     */
    context?: BaseComponentContext;
}

export interface ICustomComponentState {
  me: any;
}

export interface IMyCustomComponentWebComponentProps {
    
      /**
      * A sample string param
      */
      myStringParam?: string;
  
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
      
      /**
       * A BaseComponentContext object
      */
      context?: any;
}

export class CustomComponent extends React.Component<ICustomComponentProps, ICustomComponentState> {
    private _asyncRequest: any;

    state = {
      me: null as any,
    };

    componentDidMount() {
      this.props.context?.msGraphClientFactory.getClient("3").then(msGraphClient => {
        this._asyncRequest = msGraphClient.api('/me').get().then(me => {
            this._asyncRequest = null;
            this.setState({me});
          }
        );
      });
    }

    componentWillUnmount() {
      if (this._asyncRequest) {
        this._asyncRequest.cancel();
      }
    }

    public render() {

        // Parse custom object
        const myObject: IObjectParam = this.props.myObjectParam;

        let myName = this.state.me ? this.state.me.displayName : "Loading user...";
    
        return <div>
            {myName}<br/>
            {this.props.myStringParam} {myObject.myProperty}
        </div>;
    }
}

export class MyCustomComponentWebComponent extends BaseWebComponent {

    public constructor() {
        super();
    }

    public async connectedCallback() {

        let props = this.resolveAttributes() as IMyCustomComponentWebComponentProps;

        props.context = {}
        props.context.serviceScope = this._serviceScope;
        props.context.pageContext = this._serviceScope.consume(PageContext.serviceKey);
        props.context.spHttpClient = this._serviceScope.consume(SPHttpClient.serviceKey);  
        props.context.aadTokenProviderFactory = this._serviceScope.consume(AadTokenProviderFactory.serviceKey);
        props.context.msGraphClientFactory = this._serviceScope.consume(MSGraphClientFactory.serviceKey);

        const customComponent = <CustomComponent {...props} />;
        ReactDOM.render(customComponent, this);
    }    

    protected onDispose(): void {
        ReactDOM.unmountComponentAtNode(this);
    }
}