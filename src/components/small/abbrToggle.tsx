
import * as React from "react";
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';


export function AbbrToggle(props: { defaultValue: boolean, onAbbrText: string, offAbbrText: string, onChanged: (value: boolean) => void }) {
    return (
        <abbr title={props.defaultValue ? props.onAbbrText : props.offAbbrText}>
            <Toggle
                defaultChecked = {props.defaultValue}
                onText="Yes"
                offText="No"
                onChanged = {props.onChanged}
            />
        </abbr>
    );
}
