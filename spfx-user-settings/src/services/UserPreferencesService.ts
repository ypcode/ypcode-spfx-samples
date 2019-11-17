import { ServiceScope, ServiceKey } from "@microsoft/sp-core-library";

export interface IUserPreferences {
    favoriteSuperHero: string;
}

export interface IUserPreferencesService extends IUserPreferences {
    configure(instanceId: string): void;
}

export class UserPreferencesService implements IUserPreferencesService {
    private _instanceId: string = null;
    private _internalData: IUserPreferences = null;

    constructor(private serviceScope: ServiceScope) {

    }

    public configure(instanceId: string): void {
        this._instanceId = instanceId;
    }

    public get favoriteSuperHero(): string {
        this._ensureLoadWebPartUserPreferences();
        return this._internalData.favoriteSuperHero;
    }

    public set favoriteSuperHero(value: string) {
        this._internalData.favoriteSuperHero = value;
        this._saveWebPartUserPreferences();
    }

    private get userPreferencesKey(): string {
        return `USER_PREFS_${this._instanceId}`;
    }

    private _saveWebPartUserPreferences() {
        if (!localStorage) {
            console.error("local storage feature is not supported in this browser");
            return;
        }

        const toSave = this._internalData || { favoriteSuperHero: null };
        const userPreferencesAsString = JSON.stringify(toSave);
        localStorage.setItem(this.userPreferencesKey, userPreferencesAsString);
    }

    private _ensureLoadWebPartUserPreferences() {
        if (!localStorage) {
            console.error("local storage feature is not supported in this browser");
            return;
        }

        if (!this._internalData) {
            const userPreferencesAsString = localStorage.getItem(this.userPreferencesKey);
            if (userPreferencesAsString) {
                this._internalData = JSON.parse(userPreferencesAsString) as IUserPreferences;
            } else {
                this._internalData = {
                    favoriteSuperHero: null
                };
            }
        }
    }
}

export const UserPreferencesServiceKey = ServiceKey.create<IUserPreferencesService>("ypcode-user-preferences", UserPreferencesService);