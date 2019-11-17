import * as React from 'react';
import styles from './UserSettingsSample.module.scss';
import { IUserSettingsSampleProps } from './IUserSettingsSampleProps';
import { ISuperHero, SUPER_HEROES } from '../data/superheroes';
import { UserPreferencesServiceKey, IUserPreferences } from '../../../services/UserPreferencesService';

export interface IUserSettingsSampleState {
  favorite: string;
}

export default class UserSettingsSample extends React.Component<IUserSettingsSampleProps, IUserSettingsSampleState> {

  private userPreferences: IUserPreferences = null;
  constructor(props: IUserSettingsSampleProps) {
    super(props);
    this.userPreferences = this.props.serviceScope.consume(UserPreferencesServiceKey);
    this.state = {
      favorite: this.userPreferences.favoriteSuperHero
    };
  }

  private _isFavorite(superHero: ISuperHero): boolean {
    return this.userPreferences.favoriteSuperHero == superHero.name;
  }

  private _toggleFavorite(superHero: ISuperHero) {
    console.log("Toggle favorite super hero: ", superHero);
    if (this._isFavorite(superHero)) {
      this.userPreferences.favoriteSuperHero = null;
    } else {
      this.userPreferences.favoriteSuperHero = superHero.name;
    }
    this.setState({ favorite: this.userPreferences.favoriteSuperHero });
  }

  private _renderSuperHero(key: string, superHero: ISuperHero): JSX.Element {
    return <div key={key} className={`${styles.superHero} ${this._isFavorite(superHero) ? styles.favorite : ""}`} onClick={() => this._toggleFavorite(superHero)}>
      <img className={styles.picture} src={`${superHero.picture}`} />
      <p className={styles.title}>{superHero.name}</p>
    </div>;
  }

  public render(): React.ReactElement<IUserSettingsSampleProps> {

    return (
      <div className={styles.userSettingsSample}>
        <div className={styles.container}>
          <h1>What is your favorite super hero ?</h1>
          {SUPER_HEROES.map((superHero, ndx) => this._renderSuperHero(`super_hero_${ndx}`, superHero))}
        </div>
      </div>
    );
  }
}
