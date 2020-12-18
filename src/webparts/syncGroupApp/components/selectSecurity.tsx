import * as React from 'react';
import styles from './SyncGroupApp.module.scss';
import { sp } from "@pnp/sp";
import GroupActionAdd from "./groupActionAdd"


function SelectSecurity(props) {

    const [securityGroups, setSecurityGroups] = React.useState([{"Title": "Select a group", "ID":"0"}]); 
    const [securityGroup, setSecurityGroup] = React.useState("0");
    React.useEffect(() =>{
        getItems();
    }, [])

    React.useEffect(() => {
        console.log(securityGroup);
    })

    function getItems() {
        sp.web.lists.getByTitle("Security Groups").items.get().then(items => {
            items.forEach(item => {
               setSecurityGroups(securityGroups => securityGroups.concat({"Title": item.Title, "ID": item.GroupID}));
            });
          
           
        })
    }

    return (
        <div>
            Your group has no linked security group, you can choose one in the list : 
            <select  value={securityGroup} onChange={e => setSecurityGroup(e.currentTarget.value)}>
            {securityGroups.map(element => <option key={element.ID} value={element.ID}>{element.Title}</option>)}
            </select>
            <GroupActionAdd ID={securityGroup} context={props.context} group={props.group}/>
        </div>
    )


}

export default SelectSecurity;