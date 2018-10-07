import {ISpUser, IUserSuggestion} from "../../../../models/index"
import { IPersonaSharedProps, Persona, PersonaSize } from 'office-ui-fabric-react/lib/Persona';
import * as React from "react"
import { IFacepileProps, Facepile, OverflowButtonType, IFacepilePersona } from 'office-ui-fabric-react/lib/Facepile';
import styles from "./smallComponents.module.scss"

function _getTHumbnailUrl(email:string){
    return `https://outlook.office365.com/owa/service.svc/s/GetPersonaPhoto?email=${email}&UA=0&size=HR64x64`
}

export function SpUserPersona(props: { user: IUserSuggestion}){
    return (
        <Persona
            className = {styles.spUserPersona}
            text={props.user.Title}
            size={PersonaSize.small}
            imageUrl={_getTHumbnailUrl(props.user.Email)}
            onRenderSecondaryText={() => {
                return <a href={props.user.Email}>{props.user.Email}</a>
            }} />
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