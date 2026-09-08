import * as React from 'react';
import styles from './Slick.module.scss';
import type { ISlickProps } from './ISlickProps';
import { escape } from '@microsoft/sp-lodash-subset';

const Slick: React.FC<ISlickProps> = (props) => {

  const {
    description
  } = props;

  return (
    <section className={`${styles.slick}`}>
      <div className={styles.welcome}>
        <div>Web part property value: <strong>{escape(description)}</strong></div>
        <div>Updated 8 September 2026</div>
      </div>
    </section>
  );
}

export default Slick;
