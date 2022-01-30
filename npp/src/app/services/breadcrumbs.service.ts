import { Injectable } from '@angular/core';
import { ActivatedRouteSnapshot, Data, NavigationEnd, Router } from '@angular/router';
import { BehaviorSubject } from 'rxjs';
import { filter } from 'rxjs/operators';

export interface Breadcrumb {
  label: string;
  url: string | null;
}

@Injectable({
  providedIn: 'root'
})
export class BreadcrumbsService {

  private breadcrumbsList: Breadcrumb[] = []; 
  // Subject emitting the breadcrumb hierarchy 
  private readonly _breadcrumbs$ = new BehaviorSubject<Breadcrumb[]>([]);
  
  // Observable exposing the breadcrumb hierarchy 
  readonly breadcrumbs$ = this._breadcrumbs$.asObservable(); 

  constructor(private router: Router) { 
    this.router.events.pipe( 
      filter((event) => event instanceof NavigationEnd) 
    ).subscribe(event => { 
      this.breadcrumbsList = [];
      // Construct the breadcrumb hierarchy 
      const root = this.router.routerState.snapshot.root; 
      // const breadcrumbs: Breadcrumb[] = []; 
      this.addBreadcrumb(root, []); 
 
      // Emit the new hierarchy 
      this._breadcrumbs$.next(this.breadcrumbsList); 
    }); 
  } 
 
  private addBreadcrumb(route: ActivatedRouteSnapshot, parentUrl: string[]) { 
    if (route) { 
      // Construct the route URL 
      const routeUrl = parentUrl.concat(route.url.map(url => url.path)); 
 
      // Add an element for the current route part 
      if (route.data.breadcrumb) { 
        if (this.breadcrumbsList.length == 0) {
          const breadcrumb = { 
            label: 'Home', 
            url: '/' 
          }; 
          this.breadcrumbsList.push(breadcrumb); 
        }

        const breadcrumb = this.createBreadcrumb(route.data, routeUrl);
        if (breadcrumb) {
          this.breadcrumbsList.push(breadcrumb);
        }
      } 
      console.log('breadcrumb', this.breadcrumbsList);
      // Add another element for the next route part 
      if (route.firstChild) {
        this.addBreadcrumb(route.firstChild, routeUrl); 
      }
    } 
  } 

  public addBreadcrumbLevel(text: string, url: string | null = null) {
    const breadcrumb = { 
      label: text, 
      url
    }; 
    this.breadcrumbsList.push(breadcrumb);
    console.log('breadrumb add', this.breadcrumbsList);
    this._breadcrumbs$.next(this.breadcrumbsList); 
  }
 
  private createBreadcrumb(data: Data, routeUrl: string[]): Breadcrumb | null {
    if (typeof data.breadcrumb === 'object') {
      return { 
        label: data.breadcrumb.label, 
        url: data.breadcrumb.url 
      }; 
    } else if (typeof data.breadcrumb === 'string') {
      return { 
        label: data.breadcrumb, 
        url: '/' + routeUrl.join('/') 
      }; 
    }
    return null;
  }
}
