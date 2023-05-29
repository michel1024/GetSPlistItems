import { ISPList } from './GetSpListItemsWebPart';  
  
export default class MockHttpClient {  
    private static _items: ISPList[] = [{ 
        Title: 'E123', 
        Descripcion: 'John', 
        codigo: 'SharePoint',
        Estado: 'India',
        Fecha: new Date("2023-05-29") },];
    public static get(restUrl: string, options?: any): Promise<ISPList[]> {  
      return new Promise<ISPList[]>((resolve) => {  
            resolve(MockHttpClient._items);  
        });  
    }  
} 