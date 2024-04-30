import * as React from 'react';
import styles from './SpFxPnPjsDemo.module.scss';
import { ISpFxPnPjsDemoProps } from './ISpFxPnPjsDemoProps';
import { escape } from '@microsoft/sp-lodash-subset';
// import { getSP } from '../../../pnpjsConfig';
import { SPFI, SPFx, spfi } from '@pnp/sp';
import "@pnp/sp/webs"
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/items/get-all";
// import { GraphFI } from '@pnp/graph';
import { useState, useEffect } from 'react';

const SpFxPnPjsDemo: React.FC<ISpFxPnPjsDemoProps> = (props) => {
  // const _sp: SPFI = getSP();
  // let _graph: GraphFI;
  // _graph = getGraph();
  const {
    isDarkTheme,
    environmentMessage,
    hasTeamsContext,
    userDisplayName,
  } = props;
  const [lists, setLists] = useState<string[]>([]);


  const readList = async (): Promise<void> => {
    const sp: SPFI = spfi().using(SPFx(props.context));
    //console.log("inside useEffect")
    try {
      const response: string[] = await sp.web.lists();
      console.log(`Lists: ${response}`);
      setLists(response);
    } catch (error) {
      console.log(error)
    }
  }
  useEffect(() => {
    readList();
  }, [])
  return (
    <section className={`${styles.spFxPnPjsDemo} ${hasTeamsContext ? styles.teams : ''}`}>
      <div className={styles.welcome}>
        <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
        <h2>Well done, {escape(userDisplayName)}!</h2>
        <div>{environmentMessage}</div>
        <div>SharePoint lists in this site: {lists.map((list) => console.log(list))}</div>
      </div>
    </section>
  );
}

export default SpFxPnPjsDemo
