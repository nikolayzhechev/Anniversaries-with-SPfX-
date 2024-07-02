import { IDateTimeFieldValue } from "@pnp/spfx-property-controls/lib/PropertyFieldDateTimePicker";
import { IColor } from '@fluentui/react';

export interface IHiredDateWebPartProps {
  title: string;
  description: string;
  color: IColor;
}

export interface IHiredTime {
  context: any;
  title: any;
  color: IColor;
}

export interface IListItem {
  Id: number;
  User: {
    DisplayName: string,
    Email: string,
    HiredDate: any | Date,
    Avatar: any,
    Anniversary: number
  }
}

export interface IListItemsState {
  items: IListItem;
}

export interface IPropertyControlsTestWebPartProps {
  datetime: IDateTimeFieldValue;
}

export interface IUserProps {
  item: IListItem;
  openEditDialog: any;
  deleteItem: any;
}

export interface IpopUp {
  context: any;
  currentlySelectedId: number | any;
  setReloadData: React.Dispatch<React.SetStateAction<boolean>>;
  reloadData: boolean;
  addUserPopUpHidden: boolean;
  setAddUserPopUpHidden: React.Dispatch<React.SetStateAction<boolean>>;
  setEditUserPopUpHidden: React.Dispatch<React.SetStateAction<boolean>>;
  editUserPopUpHidden: boolean;
  editUserId: number | any;
  setEditUserId: React.Dispatch<any>;
  setEditUserHiredDate: React.Dispatch<React.SetStateAction<Date>>;
  editUserHiredDate: Date;
  editUserEmail: string[] | undefined;
}