import { ICityService } from "./ICityService";
import { ICity } from "../interfaces/ICity";
import { ServiceKey, ServiceScope } from '@microsoft/sp-core-library';
import { PageContext } from '@microsoft/sp-page-context';
import pnp, { List, ItemAddResult, ItemUpdateResult } from "sp-pnp-js";

const CITIES_LIST_NAME: string = "Cities";

export class CityService implements ICityService {
    public static readonly serviceKey: ServiceKey<ICityService> = ServiceKey.create<ICityService>('cc:ICityService', CityService);
    private _pageContext: PageContext;
    private _currentWebUrl: string;

    constructor(serviceScope: ServiceScope) {
        serviceScope.whenFinished(() => {
            this._pageContext = serviceScope.consume(PageContext.serviceKey);
            this._currentWebUrl = this._pageContext.web.absoluteUrl;

            //Setup pnp-js to work with the current web url
            pnp.setup({
                sp: {
                    baseUrl: this._currentWebUrl
                }
            });
        });
    }

    public async addCity(city: ICity): Promise<boolean> {
        return pnp.sp.web.lists.getByTitle(CITIES_LIST_NAME).items.add({
            'Title': city.Title
        }).then(async (result: ItemAddResult): Promise<boolean> => {
            let addedItem: ICity = result.data;
            console.log(addedItem);
            return pnp.sp.web.lists.getByTitle(CITIES_LIST_NAME).items.getById(addedItem.Id).update({
                Region: {
                    __metadata: { "type": "SP.Taxonomy.TaxonomyFieldValue" },
                    Label: city.Region.Label,
                    TermGuid: city.Region.TermGuid,
                    WssId: -1
                }
            }).then(async (result: ItemUpdateResult) => {
                console.log(result);
                return true;
            }, (error: any): boolean => {
                return false;
            });
        }, (error: any): boolean => {
            return false;
        });
    }

    public async updateCity(city: ICity): Promise<boolean> {
        return pnp.sp.web.lists.getByTitle(CITIES_LIST_NAME).items.getById(city.Id).update({
            Region: {
                __metadata: { "type": "SP.Taxonomy.TaxonomyFieldValue" },
                Label: city.Region.Label,
                TermGuid: city.Region.TermGuid,
                WssId: -1
            }
        }).then(async (result: ItemUpdateResult) => {
            console.log(result);
            return true;
        }, (error: any): boolean => {
            return false;
        });
    }

    public async getCity(title: string): Promise<ICity> {
        return pnp.sp.web.lists.getByTitle(CITIES_LIST_NAME)
            .items
            .select(
            "Id",
            "Title",
            "Region",
            "TaxCatchAll/Term"
            )
            .expand("TaxCatchAll")
            .filter(`Title eq '${title}'`)
            .get()
            .then((cities: any) => {
                console.log(cities[0]);
                let cityToReturn: ICity = {
                    Id: cities[0].Id,
                    Title: cities[0].Title,
                    Region:{
                        Label: cities[0].TaxCatchAll[0].Term,
                        TermGuid: cities[0].Region.TermGuid
                    }
                }
                console.log(cityToReturn);
                return cityToReturn;
            })
            .catch((error) => {
                return {
                    Title: "",
                    Region: {
                        Label: "",
                        TermGuid: ""
                    }
                };
            });
    }
}