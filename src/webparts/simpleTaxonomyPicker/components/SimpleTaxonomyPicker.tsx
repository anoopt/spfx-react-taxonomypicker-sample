import * as React from 'react';
import styles from './SimpleTaxonomyPicker.module.scss';
import { ISimpleTaxonomyPickerProps } from './ISimpleTaxonomyPickerProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { ICreateSiteState } from './ISimpleTaxonomyPickerState';
import { ICityService, CityService } from '../../../services';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Spinner, SpinnerSize, Icon } from 'office-ui-fabric-react';
import { Pivot, PivotItem } from 'office-ui-fabric-react/lib/Pivot';
import { Label } from 'office-ui-fabric-react/lib/Label';
import TaxonomyPicker, { ITaxonomyValue } from "react-taxonomypicker";
import { DefaultButton, PrimaryButton } from "office-ui-fabric-react/lib/Button";
import { ICity } from '../../../interfaces/ICity';
import "react-taxonomypicker/dist/React.TaxonomyPicker.css";

export default class SimpleTaxonomyPicker extends React.Component<ISimpleTaxonomyPickerProps, ICreateSiteState> {

  private _cityServiceInstance: ICityService;

  constructor(props: ISimpleTaxonomyPickerProps) {
    super(props);
    this.state = {
      loadingScripts: true,
      errors: [],
      status: <span></span>,
      cityInContext: {
        Title: "",
        Region: {
          Label: "",
          TermGuid: ""
        }
      },
      regionTaxValue: {
        label: "",
        path: "",
        value: ""
      },
      showUpdateControls: false
    }
  }

  public componentDidMount(): void {
    this._loadSPJSOMScripts();
  }

  private _loadSPJSOMScripts() {
    const siteColUrl = this.props.context.pageContext.site.absoluteUrl;
    try {
      SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/init.js', {
        globalExportsName: '$_global_init'
      })
        .then((): Promise<{}> => {
          return SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/MicrosoftAjax.js', {
            globalExportsName: 'Sys'
          });
        })
        .then((): Promise<{}> => {
          return SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/SP.Runtime.js', {
            globalExportsName: 'SP'
          });
        })
        .then((): Promise<{}> => {
          return SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/SP.js', {
            globalExportsName: 'SP'
          });
        })
        .then((): Promise<{}> => {
          return SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/SP.taxonomy.js', {
            globalExportsName: 'SP'
          });
        })
        .then((): void => {
          this.setState({ loadingScripts: false });
        })
        .catch((reason: any) => {
          this.setState({ loadingScripts: false, errors: [...this.state.errors, reason] });
        });
    } catch (error) {
      this.setState({ loadingScripts: false, errors: [...this.state.errors, error] });
    }
  }

  private _setTitle(value: string): void {
    this._setValue("Title", value);
  }

  private _setUpdateTitle(value: string): void {
    this._setValue("Title", value);
    let showUpdateControls: boolean = false;
    this.setState({ ...this.state, showUpdateControls });
  }

  private _setValue(field: string, value: any): void {
    let cityInContext: ICity = this.state.cityInContext;
    cityInContext[field] = value;
    this.setState({ ...this.state, cityInContext });
  }

  private _setRegion = (name, value): void => {
    console.log(value);
    if (value !== null && value !== undefined) {
      let regionLabel: string = value.label.toString();
      let regionTermGuid: string = value.value.toString();
      let cityInContext: ICity = this.state.cityInContext;
      cityInContext.Region.Label = regionLabel;
      cityInContext.Region.TermGuid = regionTermGuid;
      this.setState({ ...this.state, cityInContext });
    }
  }

  private async _getCity(): Promise<void> {
    let status: JSX.Element = <Spinner size={SpinnerSize.small} />;
    this.setState({ ...this.state, status });

    let cityInContext: ICity = await this._cityServiceInstance.getCity(this.state.cityInContext.Title);
    let regionTaxValue: ITaxonomyValue = { label: cityInContext.Region.Label, value: cityInContext.Region.TermGuid, path: "" };
    let showUpdateControls: boolean = true;
    this.setState({ ...this.state, cityInContext, regionTaxValue, showUpdateControls });

    status = <span></span>;
    this.setState({ ...this.state, status });
  }

  private async _addCity(): Promise<void> {
    let status: JSX.Element = <Spinner size={SpinnerSize.small} />;
    this.setState({ ...this.state, status });
    let cityToAdd: ICity = {
      Title: this.state.cityInContext.Title,
      Region: this.state.cityInContext.Region
    }
    let result: boolean = await this._cityServiceInstance.addCity(cityToAdd);

    let showUpdateControls: boolean = false;
    status = <span>Done!</span>;
    this.setState({ ...this.state, status, showUpdateControls });
  }

  private async _updateCity(): Promise<void> {
    let status: JSX.Element = <Spinner size={SpinnerSize.small} />;
    this.setState({ ...this.state, status });
    let cityToUpdate: ICity = {
      Id: this.state.cityInContext.Id,
      Title: this.state.cityInContext.Title,
      Region: this.state.cityInContext.Region
    }
    let result: boolean = await this._cityServiceInstance.updateCity(cityToUpdate);

    let showUpdateControls: boolean = false;
    status = <span>Done!</span>;
    this.setState({ ...this.state, status, showUpdateControls });
  }

  public render(): React.ReactElement<ISimpleTaxonomyPickerProps> {
    this._cityServiceInstance = this.props.context.serviceScope.consume(CityService.serviceKey);

    return (
      <div className={styles.simpleTaxonomyPicker}>
        <div className={styles.container}>
          <Label className={styles.title}><Icon iconName='AlignLeft'/> React Taxonomy Picker Sample</Label>
          <Pivot>
            <PivotItem linkText='Create' itemIcon='Add'>
              <div className={styles.row}>
                <div className={styles.column}>
                  <TextField label='Title' onChanged={this._setTitle.bind(this)} required={true} />
                </div>
              </div>
              {this.state.loadingScripts === false ?
                <div className={styles.row}>
                  <div className={styles.column}>
                    <Label>Region</Label>
                    <TaxonomyPicker
                      name=""
                      displayName=""
                      termSetGuid="59525237-1ba5-4f63-b967-1520f32adb6d"
                      termSetName="Regions"
                      multi={false}
                      onPickerChange={this._setRegion.bind(this)}
                    />
                  </div>
                </div> : null}
              <div className={styles.row}>
                <div className={styles.column}>
                  <PrimaryButton data-id="btnAddCity"
                    title="Add City"
                    text="Add City"
                    iconProps={ { iconName: 'Add' } }
                    onClick={this._addCity.bind(this)}
                  />
                  <div className={styles.ccStatus}>
                    {this.state.status}
                  </div>
                </div>
              </div>
            </PivotItem>
            <PivotItem linkText='Update' itemIcon='Edit'>
              <div className={styles.row}>
                <div className={styles.column}>
                  <TextField label='Title' onChanged={this._setUpdateTitle.bind(this)} />
                </div>
              </div>
              <div className={styles.row}>
                <div className={styles.column}>
                  <PrimaryButton data-id="btnGetCity"
                    title="Get City"
                    text="Get City"
                    iconProps={ { iconName: 'Search' } }
                    onClick={this._getCity.bind(this)}
                  />
                  <div className={styles.ccStatus}>
                    {this.state.status}
                  </div>
                </div>
              </div>
              {this.state.loadingScripts === false && this.state.showUpdateControls === true ?
                <div>
                  <div className={styles.row}>
                    <div className={styles.column}>
                      <TextField label='Title' onChanged={this._setTitle.bind(this)} required={true} value={this.state.cityInContext.Title} />
                    </div>
                  </div>
                  <div className={styles.row}>
                    <div className={styles.column}>
                      <Label>Region</Label>
                      <TaxonomyPicker
                        name=""
                        displayName=""
                        termSetGuid="59525237-1ba5-4f63-b967-1520f32adb6d"
                        termSetName="Regions"
                        multi={false}
                        defaultValue={this.state.regionTaxValue}
                        onPickerChange={this._setRegion.bind(this)}
                      />
                    </div>
                  </div>
                  <div className={styles.row}>
                    <div className={styles.column}>
                      <PrimaryButton data-id="btnUpdateCity"
                        title="Update City"
                        text="Update City"
                        iconProps={ { iconName: 'Edit' } }
                        onClick={this._updateCity.bind(this)}
                      />
                    </div>
                  </div>
                </div> : null}
            </PivotItem>
          </Pivot>
        </div>
      </div>
    );
  }
}
