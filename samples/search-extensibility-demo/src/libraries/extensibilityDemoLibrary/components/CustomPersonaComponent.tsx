import * as React from 'react';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import * as ReactDOM from 'react-dom';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { Persona, IPersonaProps, IPersonaSharedProps, getInitials, Icon, Link } from 'office-ui-fabric-react';
import { UrlHelper } from '../../../helpers/UrlHelper';
import * as DOMPurify from 'dompurify';
import { ITheme } from '@uifabric/styling';
import { DomPurifyHelper } from '../../../helpers/DomPurifyHelper';
import { isEmpty } from '@microsoft/sp-lodash-subset';

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
     * url to page where user will end up on click
     */
    profilePageUrl?: string;

    /**
     * The Handlebars context to inject in slide content (ex: @root)
     */
    context?: string;
}

export interface ICustomPersonaComponenState {
}

export class CustomPersonaComponent extends React.Component<ICustomPersonaComponentProps, ICustomPersonaComponenState> {

    private _domPurify: any;

    public constructor(props: ICustomPersonaComponentProps) {
        super(props);

        this._domPurify = DOMPurify.default;

        this._domPurify.setConfig({
            FORBID_TAGS: ['style'],
            WHOLE_DOCUMENT: true
        });

        this._domPurify.addHook('uponSanitizeElement', DomPurifyHelper.allowCustomComponentsHook);
        this._domPurify.addHook('uponSanitizeAttribute', DomPurifyHelper.allowCustomAttributesHook);

    }

    public render() {

        let processedProps: ICustomPersonaComponentProps = this.props;

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
            text: processedProps.userDisplayName, // This is to get the correct color for coin (used internally by the Persona component)
            onRenderInitials: (props: IPersonaSharedProps) => {

                let imageInitials = undefined;
                if (!isEmpty(processedProps.userDisplayName)) {
                    imageInitials = getInitials(UrlHelper.decode(processedProps.userDisplayName), false, false);
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

    public async connectedCallback() {
        let props = this.resolveAttributes();
        const personaItem = <CustomPersonaComponent {...props} />;
        ReactDOM.render(personaItem, this);
    }
}