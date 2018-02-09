import { ICity } from "../interfaces/ICity";

export interface ICityService {
    addCity(City: ICity): Promise<boolean>;
    updateCity(City: ICity): Promise<boolean>;
    getCity(title: string): Promise<ICity>;
}