import * as React from 'react';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import * as ReactDOM from 'react-dom';
import { PersonCard } from '@microsoft/mgt-react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';

export interface ICustomPersonCardComponentProps {
    spHttpClient: SPHttpClient;
    pageContext: PageContext;
    assistantEmail?: string;
    pronouns?: string;
}

export interface ICustomPersonCardComponentState {
    resolvedUserName: string;
}

export class CustomPersonCardComponent extends React.Component<ICustomPersonCardComponentProps, ICustomPersonCardComponentState> {
    public constructor(props: ICustomPersonCardComponentProps) {
        super(props);
        this.state = {
            resolvedUserName: ""
        };
    }

    public async componentDidMount(): Promise<void> {
        if (this.props.assistantEmail) {
            await this.getUserName(this.props.assistantEmail);
        }
    }

    public render(): React.ReactElement<ICustomPersonCardComponentProps> {

        const AdditionalDetails = (): JSX.Element => {
            return (this.props.assistantEmail || this.props.pronouns) ? <div className='additional-details'>
                {this.props.pronouns &&
                    <div className="section">
                        <div className="section__header">
                            <div className="section__title">Pronouns</div>
                        </div>
                        <div className="section__content" title="Pronouns">
                            {this.props.pronouns}
                        </div>
                    </div>
                }
                {this.state.resolvedUserName !== "" &&
                    <div className="section">
                        <div className="section__header">
                            <div className="section__title">Assistant</div>
                        </div>
                        <div className="section__content" title="Assistant">
                            {this.state.resolvedUserName}
                        </div>
                    </div>
                }
            </div> : null;
        };

        return (
            <>
                <PersonCard inheritDetails={true}>
                    {(this.props.assistantEmail || this.props.pronouns) &&
                        <AdditionalDetails />
                    }
                </PersonCard>
            </>
        );
    }

    /**
     * Gets person's display name based on the email address provided
     * @param emailAddress - person's email address
     */
    private async getUserName(emailAddress: string): Promise<void> {
        let displayName = "";
        try {
            const response: SPHttpClientResponse = await this.props.spHttpClient.post(
                this.props.pageContext.site.absoluteUrl + "/_api/web/ensureuser",
                SPHttpClient.configurations.v1,
                {
                    headers: {
                        'Accept': 'application/json;odata=nometadata',
                        'Content-type': 'application/json;odata=nometadata',
                        'odata-version': ''
                    },
                    body: JSON.stringify({ 'logonName': emailAddress })
                }
            );
            if (response.ok) {
                const responseJSON = await response.json();
                displayName = responseJSON.Title;
            }
        } catch (error) {
            console.warn('CustomPersonCardComponent: Error resolving user', error);
        }
        this.setState({ resolvedUserName: displayName });
    }
}

export class CustomPersonCardWebComponent extends BaseWebComponent {

    private _spHttpClient: SPHttpClient;
    private _pageContext: PageContext;

    constructor() {
        super();
        this._serviceScope.whenFinished(() => {
            this._spHttpClient = this._serviceScope.consume(SPHttpClient.serviceKey);
            this._pageContext = this._serviceScope.consume(PageContext.serviceKey);
        });
    }

    public connectedCallback(): void {
        const props = this.resolveAttributes();
        const personaItem = <CustomPersonCardComponent {...props} spHttpClient={this._spHttpClient} pageContext={this._pageContext} />;
        ReactDOM.render(personaItem, this);
    }
}
