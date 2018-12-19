import pnp from 'sp-pnp-js';
import { Web }  from 'sp-pnp-js/lib/sharepoint/webs';
import { model } from './../model/TopLevelMenu';
import { components } from './../components/FlyoutColumn';
import { sampleData } from './MegaMenuSampleData';

export class TheZoneGlobalNavigationService{

  private static readonly useSampleData: boolean = false;
  private static readonly levelOneListName: string = "Mega Menu - Level 1";
  private static readonly cacheKey: string = "TheZoneGlobalNavigationCache";

  public static getMenuItems(siteCollectionUrl:string): Promise<model.TopLevelMenu[]>{
    if(!TheZoneGlobalNavigationService.useSampleData){
      return new Promise<model.TopLevelMenu[]>((resolve, reject) => {
        if (!TheZoneGlobalNavigationService.useSampleData) {
          return new Promise<model.topTopLevelMenu[]>((resolve, reject) => {
              // See if we've cached the result previously.
              var topLevelItems: model.TopLevelMenu[] = pnp.storage.session.get(TheZoneGlobalNavigationService.cacheKey);
              if (topLevelItems) {
                  console.log("Found mega menu items in cache.");
                  resolve(topLevelItems);
              }
              else {
                  console.log("Didn't find mega menu items in cache, getting from list.");
                  var level1ItemsPromise = TheZoneGlobalNavigationService.getMenuItemsFromSP(TheZoneGlobalNavigationService.levelOneListName, siteCollectionUrl);
                  Promise.all([level1ItemsPromise])
                      .then((results: any[][]) => {
                          topLevelItems = TheZoneGlobalNavigationService.convertItemsFromSP(results[0]);
                          // Store in session cache.
                          pnp.storage.session.put(TheZoneGlobalNavigationService.cacheKey, topLevelItems);
                          resolve(topLevelItems);
                      });
              }
          });
      }
      else {
          return new Promise<model.TopLevelMenu[]>((resolve, reject) => {
              resolve(sampleData);
          });
      }
      });
    }
  }

  private static getMenuItemsFromSP(listName: string, siteCollectionUrl: string){
    return new Promise<model.TopLevelMenu[]>((resolve, reject) => {
      let web = new Web(siteCollectionUrl);
      // TODO : Note that passing in url and using this approach is a workaround. I would have liked to just
      // call pnp.sp.site.rootWeb.lists, however when running this code on SPO modern pages, the REST call ended
      // up with a corrupt URL. However it was OK on View All Site content pages, etc.
      web.lists
          .getByTitle(listName)
          .items
          .orderBy("SortOrder")
          .get()
          .then((items: any[]) => {
              resolve(items);
          })
          .catch((error: any) => {
              reject(error);
          });
  });
  }

  private static convertItemsFromSP(level1: any[]){
    var level1Dictionary: { [id: number]: model.TopLevelMenu; } = {};
    var level2Dictionary: { [id: number]: components.FlyoutColumn; } = {};

    // Convert level 1 items and store in dictionary.
    var level1Items: model.TopLevelMenu[] = level1.map((item: any) => {
        var newItem = {
            id: item.Id,
            text: item.Title,
            columns: []
        };

        level1Dictionary[newItem.id] = newItem;

        return newItem;
    });

    var retVal: model.TopLevelMenu[] = [];

    for (let l1 of level1Items) {
        retVal.push(l1);
    }

    return retVal;

}
}


