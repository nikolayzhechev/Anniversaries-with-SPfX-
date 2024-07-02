import * as React from 'react';
import { useState } from 'react';
import { getSP } from "../../config";
import { SPFI } from "@pnp/sp";
import "@pnp/graph/groups";
import "@pnp/graph/users";
import { DefaultButton, PrimaryButton, Dialog, DialogType, DialogFooter, IPersonaProps } from '@fluentui/react';
import {
  useId,
  Toaster,
  useToastController,
  Toast,
  ToastTitle,
  ToastTrigger,
  Link
} from "@fluentui/react-components";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { DateTimePicker, DateConvention, IDateTimePickerProps } from '@pnp/spfx-controls-react/lib/dateTimePicker';
import { IpopUp } from "../interfaces";
import styles from './PopUp.module.scss';

const PopUp: React.FC<IpopUp> = ({
    context,
    currentlySelectedId,
    setReloadData,
    reloadData,
    addUserPopUpHidden,
    setAddUserPopUpHidden,
    setEditUserPopUpHidden,
    editUserPopUpHidden,
    editUserId,
    setEditUserId,
    setEditUserHiredDate,
    editUserHiredDate,
    editUserEmail
}): React.ReactElement => {
  const _sp: SPFI = getSP(context);

  const [userId, setUserId] = useState<number | any>();
  const [newUserHireDate, setNewUserHireDate] = useState<Date>(new Date());
  const [error, setError] = useState<string>();
  const [msgPopUp, setMsgPopUp] = useState<boolean>(true);
  
  const toasterId = useId("toaster");
  const { dispatchToast } = useToastController(toasterId);

  // add new item
  const handleNewUser = (userId: number | null): void => {
    setUserId(userId);
  };

  const handleNewUserBirthday = (date: Date | undefined | null): void => {
    if (date) {
      setNewUserHireDate(date);
    }
  };

  const addNewUser = async (): Promise<void> => {
    if (errorResponseHandler(userId, newUserHireDate)){
      return;
    }

    const list = _sp.web.lists.getByTitle("Users");

    try {
      await list.items.add({
        UserId: userId,
        HiredDate: newUserHireDate
      });

      setAddUserPopUpHidden(true);
      setReloadData(!reloadData);
      notify(true, "User added successfully!");
    } catch (err) {
      setError("The specified user is already added to the list! Please add another one.");
      setMsgPopUp(false);
      console.log(err);
      notify(false);
    } finally {
      setAddUserPopUpHidden(true);
    }
  };

  const getPeoplePickerItems = (items: IPersonaProps[]): void => {
    const userId: number | any = items[0].id;

    handleNewUser(userId);
  };

  const dateTimePickerProps: IDateTimePickerProps = {
    label: "User Birth Date",
    value: newUserHireDate,
    onChange: handleNewUserBirthday,
  };

  // edit item
  const handleEditUserHiredDate = (date: Date | undefined | null): void => {
    if (date) {
      setEditUserHiredDate(date);
    }
  };

  const editDateTimePickerProps: IDateTimePickerProps = {
    label: "User Birth Date",
    value: editUserHiredDate,
    onChange: handleEditUserHiredDate,
  };

  const getEditedPeoplePickerItems = (items: any): void => {
    const userId: number | any = items[0].id;
    setEditUserId(userId);
    handleEditUser(userId);
  };

  const handleEditUser = (userId: number): void => {
    setEditUserId(userId);
  };

  const editUser = async (): Promise<void> => {
    if (errorResponseHandler(editUserEmail, editUserHiredDate)){
      return;
    }

    const list = _sp.web.lists.getByTitle("Users");
    
    try {
      await list.items.getById(currentlySelectedId).update({
        UserId: editUserId,
        HiredDate: editUserHiredDate
      });

      setEditUserPopUpHidden(true);
      setReloadData(!reloadData);
      notify(true, "User edited successfully!");
    } catch (err) {
      console.log(err);
      notify(false);
    } finally {
      setEditUserPopUpHidden(true);
    }
  };

  // notifications
  const errorResponseHandler = (user: number | string[] | null | undefined, birthDate: Date | null | undefined) => {
    if (user === undefined || user === null){
      setError("User field cannot be blank.");
      setMsgPopUp(false);
      return true;
    }
    if (birthDate === undefined || birthDate === null){
      setError("Birth date field cannot be blank.");
      setMsgPopUp(false);
      return true;
    }
    if (birthDate > getCurrentDate()){
      setError("Birth date cannot be greater than todays date.")
      setMsgPopUp(false);
      return true;
    }
    return false;
  };

  const notify = (success: boolean, msg?: string) =>
  {
    dispatchToast(
        <Toast className={styles.toastBg}>
          <ToastTitle
            className={styles.toastFont}
            action={
            <ToastTrigger>
              <Link className={styles.toastFont}>Dismiss</Link>
            </ToastTrigger>
          }>
              { success ? msg : "Error when adding new user" }
          </ToastTitle>
          
        </Toast>,
        { intent: success ? "success" : "error" }
    );
  };

  const getCurrentDate = (): Date => {
    const date: Date = new Date();
    let dd = String(date.getDate()).padStart(2, '0');
    let mm = String(date.getMonth() + 1).padStart(2, '0');
    let yyyy = date.getFullYear();
    const today = new Date(mm + '/' + dd + '/' + yyyy);

    return today;
  };

  return (
    <div>
      <Dialog
        hidden={addUserPopUpHidden}
        onDismiss={() => setAddUserPopUpHidden(true)}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Add User',
        }}>
          <PeoplePicker
            context={context}
            titleText="User Name"
            personSelectionLimit={1}
            groupName={""} // Leave this blank in case you want to filter from all users
            showtooltip={true}
            required={true}
            disabled={false}
            searchTextLimit={5}
            onChange={getPeoplePickerItems}
            principalTypes={[PrincipalType.User]}
            ensureUser={true}
            resolveDelay={1000}
          />
        <DateTimePicker {...dateTimePickerProps} dateConvention={DateConvention.Date}/>
        <DialogFooter>
          <PrimaryButton text='Submit' onClick={() => addNewUser()}></PrimaryButton>
          <DefaultButton text='Cancel' onClick={() => setAddUserPopUpHidden(true)}></DefaultButton>
        </DialogFooter>
      </Dialog>
      <Dialog
        hidden={editUserPopUpHidden}
        onDismiss={() => setEditUserPopUpHidden(true)}
        dialogContentProps={{
        type: DialogType.normal,
        title: 'Edit User'
        }}>
          <PeoplePicker
            context={context}
            titleText="User Name"
            personSelectionLimit={1}
            groupName={""} // Leave this blank in case you want to filter from all users
            showtooltip={true}
            required={true}
            disabled={false}
            searchTextLimit={5}
            onChange={getEditedPeoplePickerItems}
            principalTypes={[PrincipalType.User]}
            ensureUser={true}
            resolveDelay={1000}
            defaultSelectedUsers={editUserEmail}
          />
        <DateTimePicker {...editDateTimePickerProps} dateConvention={DateConvention.Date} />
        <DialogFooter>
          <PrimaryButton
            text='Submit'
            onClick={() => editUser()}>
          </PrimaryButton>
          <DefaultButton
            text='Cancel'
            onClick={() => setEditUserPopUpHidden(true)}>
          </DefaultButton>
        </DialogFooter>
      </Dialog>
      <Dialog
        hidden={msgPopUp}
        dialogContentProps={{
          type: DialogType.close,
          title: 'Error'
        }}
        onDismiss={() => setMsgPopUp(true)}
        >
          <div>
            <h3>Message:</h3>
            <p>{error}</p>
            <DefaultButton onClick={() => setMsgPopUp(true)}>Back</DefaultButton>
          </div>
      </Dialog>
      <>
        <Toaster toasterId={toasterId} />
      </>
    </div>
  );
};

export default PopUp;