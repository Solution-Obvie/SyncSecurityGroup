import * as React from 'react';
import styles from './SyncGroupApp.module.scss';
import { sp } from "@pnp/sp";
import GroupActionAdd from "./groupActionAdd"
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';


const dropdownStyles: Partial<IDropdownStyles> = {
    dropdown: { width: 300 },
  };
const options: IDropdownOption[] = []

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
                options.push({key: item.GroupID, text: item.Title})
               setSecurityGroups(securityGroups => securityGroups.concat({"Title": item.Title, "ID": item.GroupID}));
            });
          
           
        })
    }

    function onChange(event, item){
        setSecurityGroup(item.key);
        console.log(securityGroup)
    }
    return (
        <div>
            Your group {props.group.Title} has no linked security group, you can choose one in the list : 
            {/* <select  value={securityGroup} onChange={e => setSecurityGroup(e.currentTarget.value)}>
            {securityGroups.map(element => <option key={element.ID} value={element.ID}>{element.Title}</option>)}
            </select> */}
            <div className={styles.selectContainer}>
            <Dropdown
            placeholder="Select a security Group"
            options={options}
            onChange={onChange}
            styles={dropdownStyles}
            />
            <GroupActionAdd ID={securityGroup} context={props.context} group={props.group} setGroup={props.setGroup}  setProgress={props.setProgress} progress={props.progress}/>
            </div>
         
        </div>
    )


}

export default SelectSecurity;