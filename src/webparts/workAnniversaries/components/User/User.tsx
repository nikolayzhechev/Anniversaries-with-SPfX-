import * as React from 'react';
import styles from "./User.module.scss";
import { Card, CardHeader, CardPreview } from "@fluentui/react-components";
import { Body1, Caption1 } from "@fluentui/react-components";
import { IconButton } from '@fluentui/react';
import { IUserProps } from '../interfaces';

const User: React.FC<IUserProps> = ({ item, openEditDialog, deleteItem }): React.ReactElement => {
    return (
      <div className={styles.cardContainer}>
        <Card
            className={styles.card}
            key={item.Id}
            orientation="horizontal"
            >
            <CardPreview className={styles.cardImg}>
              <img
                src={item.User.Avatar}
                alt='Avatar'>
              </img>
            </CardPreview>
            <CardHeader
              header={
                <section>
                  <div>
                    <Body1 className={styles.cardBody}>
                        <b>{item.User.DisplayName}</b>
                    </Body1>
                  </div>
                  <div>
                    <Caption1>
                      {item.User.Email}
                    </Caption1>
                  </div>
                </section>
              }
              description={
                <Body1>
                  Anniversary: <b>{item.User.Anniversary}</b>
                  Hired Date: {item.User.HiredDate}
                </Body1>
              }
            />
            <IconButton
              iconProps={{ iconName: 'Edit'}}
              onClick={() => openEditDialog(item.Id)}
            />
            <IconButton
              iconProps={{ iconName: 'Delete' }}
              onClick={() => deleteItem(item.Id)}
            />
        </Card>
      </div>
    )
};

export default User;