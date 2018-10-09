import {ISpUser, IUserSuggestion} from "../../../../models/index"
import { IPersonaProps  , Persona, PersonaSize } from 'office-ui-fabric-react/lib/Persona';
import * as React from "react"
import { IFacepileProps, Facepile, OverflowButtonType, IFacepilePersona } from 'office-ui-fabric-react/lib/Facepile';
import styles from "./smallComponents.module.scss"

function _getTHumbnailUrl(email:string){
    return `https://outlook.office365.com/owa/service.svc/s/GetPersonaPhoto?email=${email}&UA=0&size=HR64x64`
}

export function SpUserPersona(props: { user: IUserSuggestion, personaProps?: IPersonaProps}){
    let sharedProps = props.personaProps || {};
    sharedProps.text = sharedProps.text || props.user.Title;
    sharedProps.size = sharedProps.size || PersonaSize.small;
    sharedProps.imageUrl = sharedProps.imageUrl || _getTHumbnailUrl(props.user.Email);
    sharedProps.onRenderSecondaryText = sharedProps.onRenderSecondaryText || (() => {
        return <a href={props.user.Email}>{props.user.Email}</a>
    });
    
    return (
        <Persona
            className = {styles.spUserPersona}
            {...sharedProps}/>
    )
}

export function SpUsersFacepile(props:{users:Array<ISpUser>}){
    return (
        <Facepile
            personas={props.users.slice(0, 10).map<IFacepilePersona>(u => {
                return {
                    imageUrl: _getTHumbnailUrl(u.Email),
                    personaName: u.Title
                }
            })}
            overflowPersonas={props.users.slice(10).map<IFacepilePersona>(u => {
                return {
                    imageUrl: _getTHumbnailUrl(u.Email),
                    personaName: u.Title
                }
            })}
        />
    )
}