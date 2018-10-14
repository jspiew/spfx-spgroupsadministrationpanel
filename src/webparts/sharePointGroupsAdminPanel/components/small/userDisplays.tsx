import {ISpUser, IUserSuggestion} from "../../../../models/index"
import { IPersonaProps  , Persona, PersonaSize } from 'office-ui-fabric-react/lib/Persona';
import * as React from "react"
import { IFacepileProps, Facepile, OverflowButtonType, IFacepilePersona } from 'office-ui-fabric-react/lib/Facepile';
import styles from "./smallComponents.module.scss"

function _getTHumbnailUrl(email:string){
    return `https://outlook.office365.com/owa/service.svc/s/GetPersonaPhoto?email=${email}&UA=0&size=HR64x64`
}

export function SpUserPersona(props: { user: IUserSuggestion, onDelete?: (user:IUserSuggestion) => void, personaProps?: IPersonaProps}){
    let sharedProps = props.personaProps || {};
    sharedProps.text = sharedProps.text || props.user.Title;
    sharedProps.size = sharedProps.size || PersonaSize.small;
    sharedProps.imageUrl = sharedProps.imageUrl || _getTHumbnailUrl(props.user.Email);
    sharedProps.onRenderSecondaryText = sharedProps.onRenderSecondaryText || (() => {
        return <a href={props.user.Email} className={styles.secondaryTextColor}>
                    {props.user.Email &&<i className="ms-Icon ms-Icon--Mail" aria-hidden="true"></i>}
                    {props.user.Email}
                </a>
    });
    return (
        <div className = {styles.spUserPersona}>
            {props.onDelete && <i className={`${styles.deleteIcon} ms-Icon ms-Icon--Delete`} aria-hidden="true" onClick={() => { props.onDelete(props.user) }}></i>}
            <Persona
                className = {styles.persona}
                {...sharedProps}/>
        </div>
        
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