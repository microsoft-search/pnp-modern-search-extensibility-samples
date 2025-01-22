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

export interface ICustomPersonCardComponenState {
    resolvedUserName: string;
}

export class CustomPersonCardComponent extends React.Component<ICustomPersonCardComponentProps, ICustomPersonCardComponenState> {
    public constructor(props: ICustomPersonCardComponentProps) {
        super(props);
        this.state = {
            resolvedUserName: ""
        }

    }
    public async componentDidMount() {
        if (this.props.assistantEmail) {
            this.getUserName(this.props.assistantEmail);
        }
    }
    public render() {

        var AdditionalDetails = AdditionalDetails = (MgtTemplateProps) => {
            return (this.props.assistantEmail || this.props.pronouns) && <div className='additional-details'>
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
                {this.state.resolvedUserName != "" &&
                    <div className="section">
                        <div className="section__header">
                            <div className="section__title">Assistant</div>
                        </div>
                        <div className="section__content" title="Assistant">
                            {this.state.resolvedUserName}
                        </div>
                    </div>
                }
            </div>
        };

        return (
            <>
                <PersonCard inheritDetails={true}>
                    {(this.props.assistantEmail || this.props.pronouns) &&
                        <AdditionalDetails template="additional-details">
                        </AdditionalDetails>
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
        await this.props.spHttpClient.post(this.props.pageContext.site.absoluteUrl + "/_api/web/ensureuser", SPHttpClient.configurations.v1, {
            headers: {
                'Accept': 'application/json;odata=nometadata',
                'Content-type': 'application/json;odata=nometadata',
                'odata-version': ''
            },
            body: JSON.stringify({ 'logonName': emailAddress })
        }).then((response: SPHttpClientResponse) => {
            if (response.ok) {
                response.json().then((responseJSON) => {
                    console.log(responseJSON);
                    displayName = responseJSON.Title;
                    this.setState({
                        resolvedUserName: displayName
                    })
                });
            } else {
                response.json().then((responseJSON) => {
                    console.log(responseJSON);
                    this.setState({
                        resolvedUserName: displayName
                    })
                });
            }

        }).catch(error => {
            console.log(error);
            this.setState({
                resolvedUserName: displayName
            })
        });

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
    public connectedCallback() {
        let props = this.resolveAttributes();
        const personaItem = <CustomPersonCardComponent {...props} spHttpClient={this._spHttpClient} pageContext={this._pageContext} />;
        ReactDOM.render(personaItem, this);
    }
}