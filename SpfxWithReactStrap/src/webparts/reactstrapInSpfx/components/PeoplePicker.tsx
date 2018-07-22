import * as React from 'react';
import {
    PeoplePickerState,
    IClientPeoplePickerSearchUser,
    IEnsurableSharePointUser,
    IEnsureUser,
    SharePointUserPersona
} from "./PeoplePickerProperties";
import { IPersonaWithMenu } from 'office-ui-fabric-react/lib/components/pickers/PeoplePicker/PeoplePickerItems/PeoplePickerItem.types';
import {
    CompactPeoplePicker,
    IBasePickerSuggestionsProps,
    NormalPeoplePicker
} from 'office-ui-fabric-react/lib/Pickers';
import { IPersonaProps } from 'office-ui-fabric-react/lib/Persona';
import {
    assign,
    autobind
} from 'office-ui-fabric-react/lib/Utilities';
import { IContextualMenuItem } from 'office-ui-fabric-react/lib/ContextualMenu';
import {
    SPHttpClient,
    SPHttpClientResponse,
    SPHttpClientBatch
} from '@microsoft/sp-http';
import { Promise } from 'es6-promise';
import * as lodash from 'lodash';


const suggestionProps: IBasePickerSuggestionsProps = {
    suggestionsHeaderText: 'Suggested People',
    noResultsFoundText: 'No results found',
    loadingText: 'Loading'
};

export interface IPeoplePicker {
    description: string;
    spHttpClient: SPHttpClient;
    siteUrl: string;
    principalTypeUser: boolean;
    principalTypeSharePointGroup: boolean;
    principalTypeSecurityGroup: boolean;
    principalTypeDistributionList: boolean;
    numberOfItems: number;
    onChange?: (items: SharePointUserPersona[]) => void;
}

export default class PeoplePicker extends React.Component<IPeoplePicker, PeoplePickerState>{
    private _peopleList;
    private contextualMenuItems: IContextualMenuItem[] = [
        {
            key: 'newItem',
            icon: 'circlePlus',
            name: 'New'
        },
        {
            key: 'upload',
            icon: 'upload',
            name: 'Upload'
        },
        {
            key: 'divider_1',
            name: '-',
        },
        {
            key: 'rename',
            name: 'Rename'
        },
        {
            key: 'properties',
            name: 'Properties'
        },
        {
            key: 'disabled',
            name: 'Disabled item',
            disabled: true
        }
    ]
    constructor(props) {
        super(props);
        this._peopleList = [];
        this.state = {
            currentPicker: 1,
            delayResults: false,
            selectedItems: []
        };
    }

    public render(): React.ReactElement<IPeoplePicker> {
        return (
            <NormalPeoplePicker
                onChange={this._onChange.bind(this)}
                onResolveSuggestions={this._onFilterChanged}
                getTextFromItem={(persona: IPersonaProps) => persona.primaryText}
                pickerSuggestionsProps={suggestionProps}
                className={'ms-PeoplePicker'}
                key={'normal'}
            />
        );
    }

    private _onChange(items: any[]) {
        this.setState({
            selectedItems: items
        });
        if (this.props.onChange) {
            this.props.onChange(items);
        }
    }

    @autobind
    private _onFilterChanged(filterText: string, currentPersonas: IPersonaProps[], limitResults?: number) {
        if (filterText) {
            if (filterText.length > 2) {
                return this._searchPeople(filterText, this._peopleList);
            }
        } else {
            return [];
        }
    }

    private _searchPeople(terms: string, results: IPersonaProps[]): IPersonaProps[] | Promise<IPersonaProps[]> {
        const userRequestUrl: string = `${this.props.siteUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`;
        let principalType: number = 0;
        if (this.props.principalTypeUser === true) {
            principalType += 1;
        }
        if (this.props.principalTypeSharePointGroup === true) {
            principalType += 8;
        }
        if (this.props.principalTypeSecurityGroup === true) {
            principalType += 4;
        }
        if (this.props.principalTypeDistributionList === true) {
            principalType += 2;
        }
        const userQueryParams = {
            'queryParams': {
                'AllowEmailAddresses': true,
                'AllowMultipleEntities': false,
                'AllUrlZones': false,
                'MaximumEntitySuggestions': this.props.numberOfItems,
                'PrincipalSource': 15,
                // PrincipalType controls the type of entities that are returned in the results.
                // Choices are All - 15, Distribution List - 2 , Security Groups - 4, SharePoint Groups - 8, User - 1.
                // These values can be combined (example: 13 is security + SP groups + users)
                'PrincipalType': principalType,
                'QueryString': terms
            }
        };

        return new Promise<SharePointUserPersona[]>((resolve, reject) =>
            this.props.spHttpClient.post(userRequestUrl,
                SPHttpClient.configurations.v1, { body: JSON.stringify(userQueryParams) })
                .then((response: SPHttpClientResponse) => {
                    return response.json();
                })
                .then((response: { value: string }) => {
                    let userQueryResults: IClientPeoplePickerSearchUser[] = JSON.parse(response.value);
                    let persons = userQueryResults.map(p => new SharePointUserPersona(p as IEnsurableSharePointUser));
                    return persons;
                })
                .then((persons) => {
                    const batch = this.props.spHttpClient.beginBatch();
                    const ensureUserUrl = `${this.props.siteUrl}/_api/web/ensureUser`;
                    const batchPromises: Promise<IEnsureUser>[] = persons.map(p => {
                        var userQuery = JSON.stringify({ logonName: p.User.Key });
                        return batch.post(ensureUserUrl, SPHttpClientBatch.configurations.v1, {
                            body: userQuery
                        })
                            .then((response: SPHttpClientResponse) => response.json())
                            .then((json: IEnsureUser) => json);
                    });

                    var users = batch.execute().then(() => Promise.all(batchPromises).then(values => {
                        values.forEach(v => {
                            let userPersona = lodash.find(persons, o => o.User.Key == v.LoginName);
                            if (userPersona && userPersona.User) {
                                let user = userPersona.User;
                                lodash.assign(user, v);
                                userPersona.User = user;
                            }
                        });

                        resolve(persons);
                    }));
                }, (error: any): void => {
                    reject(this._peopleList = []);
                }));
    }
}