export interface ICity{
    Id?: number;
    Title: string;
    Region: {
        Label: string,
        TermGuid: string
    };
}