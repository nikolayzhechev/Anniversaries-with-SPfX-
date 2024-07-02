import * as React from 'react';
import { useEffect, useState } from 'react';
import styles from "./HiredTime.module.scss";
import { IListItem } from "../interfaces";
import { getSP, getGraph } from "../../config";
import { SPFI } from "@pnp/sp";
import "@pnp/graph/groups";
import "@pnp/graph/users";
import { PrimaryButton, DefaultButton, Dropdown, IDropdownOption } from '@fluentui/react';
import User from "../User/User";
import PopUp from "../PopUp/PopUp";
import dropDownMonthOptions from "./monts";
import { GraphFI } from '@pnp/graph';
import { IHiredDateProps } from './IHiredTimeProps';

const HiredTime: React.FC<IHiredDateProps> = ({context, title, color}): React.ReactElement => {
  const _sp: SPFI = getSP(context);
  const _graph: GraphFI = getGraph(context);
  const currentMonth: number[] = Array(1).fill(new Date().getMonth() + 1);

  const [itemsState, setItemsState] = useState<any[]>([]);
  const [addUserPopUpHidden, setAddUserPopUpHidden] = useState<boolean>(true);
  const [reloadData, setReloadData] = useState<boolean>(false);
  const [editUserPopUpHidden, setEditUserPopUpHidden] = useState<boolean>(true);
  const [editUserId, setEditUserId] = useState<number | any>();
  const [editUserHiredDate, setEditUserHiredDate] = useState<Date>(new Date());
  const [editUserEmail, setEditUserEmail] = useState<string[]>();
  const [currentlySelectedId, setCurrentlySelectedId] = useState<number | any>();
  const [isLoading, setLoading] = useState(true)
  const [selectedMonth, setSelectedMonth] = useState<number[]>([]);
  const [isDefaultMonth, setIsDefaultMonth] = useState<boolean>(true);

  useEffect(() => {
    const fetchData = async () => {
      try{
        setLoading(true);
        const response: IListItem[] = await _sp.web.lists
          .getByTitle("Users")
          .items
          .select("Id", "HiredDate", "User/EMail", "User/Title", "User/Id")
          .expand("User/EMail", "User/Title", "User/Id")();

        const filteredDataByMonth: any[] = filterResponseData(response);
        const items = [];
        
        for (const item of filteredDataByMonth) {
          if(item.User.EMail !== undefined || item.User.Id !== undefined) {
            const email: string = item.User.EMail;
            const user = await _sp.web.getUserById(item.User.Id)();
            const userProperties = await _sp.profiles.getPropertiesFor(user.LoginName);
            _graph.users;

            const anniversary: number = calcYear(new Date(item.HiredDate).getFullYear());

            const userObject: IListItem = {
              Id: item.Id,
              User: {
                DisplayName: item.User.Title,
                Email: email,
                HiredDate: item.HiredDate,
                Avatar: userProperties.PictureUrl,
                Anniversary: anniversary
              }
            }

            items.push(userObject);
          } else {
            console.log("User or email is undefined.")
          }
        };
        setItemsState(items);
      } catch(err) {
        console.log(err);
      }
      finally {
        setLoading(false);
      }
  };
    fetchData();
  }, [reloadData]);

  const calcYear = (hiredYear: number): number => {
    const currentYear: number = new Date().getFullYear();
    return currentYear - hiredYear;
  };

  const getMonth = (): number[] => {
    if (isDefaultMonth){
      const monthArray = [];
      const current = new Date();
      const month = current.getMonth() + 1;
      
      monthArray.push(month);
      
      return monthArray;
    } else {
      return selectedMonth;
    }
  };

  const filterResponseData = (response: any[]): any[] => {
    const filteredItems: any[] = [];

    for(let item of response){
      const getMonts: number[] = getMonth();
      const month: number = new Date(item.HiredDate).getMonth() + 1;

      getMonts.forEach((dropDownMonth: number) => {
        if (month === dropDownMonth){
          filteredItems.push(item);
        }
      });
    };

    const sortedItems = filteredItems.sort((a, b): any => {
      return new Date(a.HiredDate).getMonth() - new Date(b.HiredDate).getMonth() || new Date(a.HiredDate).getDate() - new Date(b.HiredDate).getDate();
    })
    .map((item): any => {
      item.HiredDate = new Date(item.HiredDate).toLocaleDateString();
      return item;
    })

    return sortedItems;
  };

  // edit item
  const openEditDialog = (id: number): void => {
    setEditUserPopUpHidden(false);

    setCurrentlySelectedId(id);

    const user: IListItem | undefined = itemsState.find((each) => each.Id === id);

    if (user) {
      const date: Date = new Date(user.User.HiredDate);
      const defaultSelectedEmailArray: string[] = [user.User.Email];
      setEditUserHiredDate(date);
      setEditUserEmail(defaultSelectedEmailArray);
    }
  };

  // delete item
  const deleteItem = async (itemid: number): Promise<void> => {
    try {
      await _sp.web.lists.getByTitle("Users").items.getById(itemid).delete();

      setReloadData(!reloadData);
    } catch (err) {
      console.log(err);
    }
  }

  const handleMonthSelection = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    setSelectedMonth(
      item.selected ?
        [...selectedMonth, item.key as number]
        : selectedMonth.filter(key => key !== item.key),
        //: dropDownMonthOptions.map(x => x.key as number)
      );

      setIsDefaultMonth(false);
      setReloadData(!reloadData);
  };

  const resetToDefaultHandler = (): void => {
    setSelectedMonth([]);
    setIsDefaultMonth(true);
    setReloadData(!reloadData);
  };

  return (
    <section>
        <h1 className={styles.title}>{title}</h1>
        <div className={styles.containerDropDown}>
          <Dropdown
            label={`Results are shown for the selected month/s:`}
            options={dropDownMonthOptions}
            placeholder='Select options'
            selectedKeys={isDefaultMonth ? currentMonth : selectedMonth}
            onChange={handleMonthSelection}
            multiSelect
            className={styles.dropDown}
          >
          </Dropdown>
          <DefaultButton
            text='Reset'
            onClick={resetToDefaultHandler}
            className={styles.dropDownResetBtn}
            >
          </DefaultButton> 
        </div>
        {
          isLoading ?
          <div>
            <p>Loading...</p>
          </div> :
          <div>
            <p className={styles.resultTxtStyle}>Results:</p>
            {
              itemsState.length > 0 ?
              itemsState.map((item: IListItem) => 
                <User
                  item={item}
                  openEditDialog={openEditDialog}
                  deleteItem={deleteItem}/>
                ) : selectedMonth.length === 0 ? 
                  <p>Please select a month from the dropdown</p>
                  : <p>There are no anniversaries in this period.</p>
            }
          </div>
        }
        <PrimaryButton
          className={styles.primaryBtn}
          text='Add User'
          onClick={() => setAddUserPopUpHidden(false)}>
        </PrimaryButton>
        <PopUp
          context={context}
          currentlySelectedId={currentlySelectedId}
          setReloadData={setReloadData}
          reloadData={reloadData}
          addUserPopUpHidden={addUserPopUpHidden}
          setAddUserPopUpHidden={setAddUserPopUpHidden}
          editUserPopUpHidden={editUserPopUpHidden}
          setEditUserPopUpHidden={setEditUserPopUpHidden}
          editUserId={editUserId}
          setEditUserId={setEditUserId}
          editUserHiredDate={editUserHiredDate}
          setEditUserHiredDate={setEditUserHiredDate}
          editUserEmail={editUserEmail}
          >
        </PopUp>
      </section>
    );
}

export default HiredTime;