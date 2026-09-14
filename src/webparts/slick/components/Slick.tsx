import * as React from 'react';
import styles from './Slick.module.scss';
import type { ISlickProps } from './ISlickProps';
import { escape } from '@microsoft/sp-lodash-subset';

const Slick: React.FC<ISlickProps> = (props) => {

  const {
    description,
    level
  } = props;

  return (
    <section className={`${styles.slick}`}>
      <div className={styles.welcome}>
        <div>Web part property value: <strong>{escape(description)}</strong></div>
        <div> Level is {props.level}</div>
        <div>Updated 13 September 2026</div>
      </div>
    </section>
  );
}

export default Slick;
