import * as React from 'react';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import * as ReactDOM from 'react-dom';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { Persona, IPersonaProps, IPersonaSharedProps, getInitials, Icon, Link } from '@fluentui/react';
import { ITheme } from '@fluentui/react';
import { isEmpty } from '@microsoft/sp-lodash-subset';
import * as DOMPurify from 'dompurify';

export interface ICustomPersonaComponentProps {

    /**
     * The item context
     */
    item?: { [key: string]: any };

    /**
     * The persona coin image URL
     */
    imageUrl?: string;

    /**
     * Persona card primary text
     */
    userDisplayName?: string;

    /**
     * Persona card secondary text
     */
    jobTitle?: string;

    /**
     * Persona card tertiary text
     */
    userEmail?: string;

    /**
     * Persona card quaternary text
     */
    office?: string;

    /**
     * Persona card quinary text
     */
    pronouns?: string;

    /**
     * The current theme settings
     */
    themeVariant?: IReadonlyTheme;

    /**
     * URL to page where user will end up on click
     */
    profilePageUrl?: string;

    /**
     * The Handlebars context to inject in slide content (ex: @root)
     */
    context?: string;
}

export interface ICustomPersonaComponentState {
}

export class CustomPersonaComponent extends React.Component<ICustomPersonaComponentProps, ICustomPersonaComponentState> {

    private _domPurify: any;

    public constructor(props: ICustomPersonaComponentProps) {
        super(props);

        this._domPurify = (DOMPurify as any).default || DOMPurify;

        this._domPurify.setConfig({
            FORBID_TAGS: ['style'],
            WHOLE_DOCUMENT: true
        });
    }

    public render(): React.ReactElement<ICustomPersonaComponentProps> {

        const processedProps: ICustomPersonaComponentProps = this.props;

        const persona: IPersonaProps = {
            theme: this.props.themeVariant as ITheme,
            imageUrl: this.props.imageUrl ? this.props.imageUrl : processedProps.imageUrl,
            imageShouldFadeIn: false,
            imageShouldStartVisible: true,
            styles: {
                root: {
                    height: '100%'
                }
            },
            text: processedProps.userDisplayName,
            onRenderInitials: (props: IPersonaSharedProps) => {

                let imageInitials: string = undefined;
                if (!isEmpty(processedProps.userDisplayName)) {
                    imageInitials = getInitials(processedProps.userDisplayName, false, false);
                }

                return imageInitials ? <span>{imageInitials}</span> : <Icon iconName="Contact" />;
            },
            onRenderPrimaryText: (props: IPersonaProps) => {
                return <><div style={{ display: 'inline', fontWeight: 'bold' }}
                    dangerouslySetInnerHTML={{ __html: this._domPurify.sanitize(processedProps.userDisplayName) }}></div>&nbsp;
                    <div className='pronouns' dangerouslySetInnerHTML={{ __html: this._domPurify.sanitize(processedProps.pronouns) }}></div></>;
            },
            onRenderSecondaryText: (props: IPersonaProps) => {
                return <div style={{ display: 'inline' }} dangerouslySetInnerHTML={{ __html: this._domPurify.sanitize(processedProps.jobTitle) }}></div>;
            },
            onRenderTertiaryText: (props: IPersonaProps) => {
                return <div style={{ display: 'inline' }} dangerouslySetInnerHTML={{ __html: this._domPurify.sanitize(processedProps.userEmail) }}></div>;
            },
            onRenderOptionalText: (props: IPersonaProps) => {
                return <div>
                    <div dangerouslySetInnerHTML={{ __html: this._domPurify.sanitize(processedProps.office) }}></div>
                </div>;
            }
        };

        return (
            <Link data-interception="off" target="_blank" href={processedProps.profilePageUrl + "?email=" + processedProps.userEmail}
                className="customPersonaLink">
                <Persona {...persona} size={15}></Persona>
            </Link>
        );
    }
}

export class CustomPersonaWebComponent extends BaseWebComponent {
    public constructor() {
        super();
    }

    public async connectedCallback(): Promise<void> {
        const props = this.resolveAttributes();
        const personaItem = <CustomPersonaComponent {...props} />;
        ReactDOM.render(personaItem, this);
    }

    protected onDispose(): void {
        ReactDOM.unmountComponentAtNode(this);
    }
}
