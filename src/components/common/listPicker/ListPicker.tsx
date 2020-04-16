import * as React from "react";
import { useEffect, useState } from "react";
import { ServiceScope } from '@microsoft/sp-core-library';
import { Dropdown, IDropdownOption } from "office-ui-fabric-react/lib/Dropdown";
import { IList } from "../../../models/IList";
import { SitesServiceKey } from "../../../services/sites/SitesService";

export interface IListPickerProps {
    label?: string;
    webUrl: string;
    serviceScope: ServiceScope;
    onListSelected: (listUrl: string) => void;
}

export const ListPicker = (props: IListPickerProps) => {

    const [availableLists, setAvailableLists] = useState<IList[]>([]);

    useEffect(() => {
        // Load the lists from the specified web
        const sitesService = props.serviceScope.consume(SitesServiceKey);
        sitesService.getSiteLists(props.webUrl).then(lists => setAvailableLists(lists));
    }, [props.webUrl]);

    const onChange = (ev: any, option: IDropdownOption) => {
        props.onListSelected(option.key as string);
    };

    return <Dropdown
        disabled={!availableLists || (availableLists && availableLists.length == 0)}
        label={props.label}
        options={availableLists.map(al => ({ key: al.url, text: al.title }))}
        onChange={onChange} />;
};