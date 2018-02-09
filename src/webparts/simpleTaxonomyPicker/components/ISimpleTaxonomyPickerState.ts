import { ICity } from "../../../interfaces/ICity";
import { ITaxonomyValue } from "react-taxonomypicker/dist/components/TaxonomyPicker/ITaxonomyPickerProps";

export interface ICreateSiteState {
    loadingScripts: boolean;
    errors?: string[];
    status: JSX.Element;
    cityInContext: ICity;
    regionTaxValue: ITaxonomyValue;
    showUpdateControls: boolean;
}