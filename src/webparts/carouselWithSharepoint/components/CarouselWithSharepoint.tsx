// import * as React from 'react';
// import styles from './CarouselWithSharepoint.module.scss';
// import { ICarouselWithSharepointProps } from './ICarouselWithSharepointProps';
// import "bootstrap/dist/css/bootstrap.css";
// import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
// import { ImageFit } from 'office-ui-fabric-react';
// import { CarouselIndicatorShape } from '@pnp/spfx-controls-react/lib/controls/carousel';
// import { ICarouselImageProps } from '@pnp/spfx-controls-react/lib/controls/carousel';
// import { Carousel } from "@pnp/spfx-controls-react/lib/Carousel";
// import { CarouselButtonsDisplay } from "@pnp/spfx-controls-react/lib/Carousel";
// import { CarouselButtonsLocation } from "@pnp/spfx-controls-react/lib/Carousel";




// //API 1
// export interface ISliderCarouselListItem {
//   Title: string;
//   Description: string;
//   BannerImageUrl: { Url: string };
//   PromotedState: number;
//   isLive : string;
//   isExpired : boolean;
//   File: {ServerRelativeUrl:string}; 
//   Created:string;
// }

// export interface ISliderCarouselDemoState {
//   value: ISliderCarouselListItem[];
// }

// //API 2
// export interface ISliderCarouselListItemCollection2 {
//   Title: string;
//   Description: string;
//   BannerImageUrl: { Url: string };
//   PromotedState: number;
//   isLive : string;
//   isExpired : boolean;
//   File: {ServerRelativeUrl:string}; 
//   Created:string;
// }

// export interface ISliderCarouselDemoStateCollection2 {
//   valueCollection2: ISliderCarouselListItemCollection2[];
// }

// //API 3
// export interface ISliderCarouselListItemCollection3 {
//   Title: string;
//   Description: string;
//   BannerImageUrl: { Url: string };
//   PromotedState: number;
//   isLive : string;
//   isExpired : boolean;
//   File: {ServerRelativeUrl:string}; 
//   Created:string;
// }

// export interface ISliderCarouselDemoStateCollection3 {
//   valueCollection3: ISliderCarouselListItemCollection3[];
// }

// export default class CarouselWithSharepoint extends React.Component<
//   ICarouselWithSharepointProps,
//   { value: ISliderCarouselListItem[],valueCollection2:ISliderCarouselListItemCollection2[],valueCollection3:ISliderCarouselListItemCollection3[]}
// > { 
//   constructor(props: ICarouselWithSharepointProps) {
//     super(props);
//     this.state = {
//       value: [],
//       valueCollection2:[],
//       valueCollection3:[]
//     };
//   }
//   private getCarouselListContent = () => {
//     try {
//       const requestUrl1 = `https://artshouselimited.sharepoint.com/sites/AHLMainHub/_api/web/lists/getbytitle('Site%20Pages')/items?$Select*,File/*&$expand=File`;
//       const requestUrl2 = `https://artshouselimited.sharepoint.com/sites/AHLHRHub/_api/web/lists/getbytitle('Site%20Pages')/items?$Select*,File/*&$expand=File`;
//       const requestUrl3 = `https://artshouselimited.sharepoint.com/sites/AHLITHub/_api/web/lists/getbytitle('Site%20Pages')/items?$Select*,File/*&$expand=File`;
//       Promise.all([
//         this.props.spHttpClient.get(requestUrl1, SPHttpClient.configurations.v1),
//         this.props.spHttpClient.get(requestUrl2, SPHttpClient.configurations.v1),
//         this.props.spHttpClient.get(requestUrl3, SPHttpClient.configurations.v1)
//       ])
//         .then(([response1, response2, response3]: [SPHttpClientResponse, SPHttpClientResponse, SPHttpClientResponse]) => {
//           if (response1.ok && response2.ok && response3.ok) {
//             return Promise.all([response1.json(), response2.json(), response3.json()]);
//           }
//         })
//         .then(([data1, data2, data3]: [{ value: any }, { value: any }, { value: any }]) => {
//           if (data1 != null && data2 != null && data3 != null) {
//             this.setState({
//               value: data1.value,
//               valueCollection2: data2.value,
//               valueCollection3: data3.value
//             });
//           }
//         });
//     } catch (error) {
//       console.log('error in service ', error);
//     }
//   };
  

//   componentDidMount = () => {
//     this.getCarouselListContent();
//   };

// //success merge 3 collection
// public render(): React.ReactElement<ICarouselWithSharepointProps> {
//   let collection = this.state.value;
//   let collection2 = this.state.valueCollection2;
//   let collection3 = this.state.valueCollection3;

//   let imageURL: string[] =[];

//   // Merge the three collections
//   let mergedCollection = collection.concat(collection2, collection3);
//   let CarouselElement: ICarouselImageProps[] = [];
//   // mergedCollection.length > 0 &&
//   //   mergedCollection
//   //     .filter(item => item.PromotedState === 2 && item.isLive === "Yes")
//   //     .slice(0, this.props.NumberOfNews === undefined || this.props.NumberOfNews <= 0 ? 10 : this.props.NumberOfNews)
//   //     .map((data, index) => {
//   //       CarouselElement.push({
//   //         imageSrc: data.BannerImageUrl.Url,
//   //         title: data.Title,
//   //         description: data.Description,
//   //         url: 'https://artshouselimited.sharepoint.com' + data.File["ServerRelativeUrl"],
//   //         showDetailsOnHover: true,
//   //         imageFit: ImageFit.cover,
//   //         imgStyle:{
//   //           height:'338px !important',
//   //           width: '100% !important',
//   //           zIndex: 0,
//   //         }
        
//   //       });
//   //     });
// if (mergedCollection.length > 0) {
//   // Sort the list items by the Created field in desending order
//   mergedCollection.sort((a, b) => new Date(b.Created).getTime() - new Date(a.Created).getTime());


//   mergedCollection
//     .filter(item => item.PromotedState === 2 && item.isLive === "Yes")
//     .slice(0, this.props.NumberOfNews === undefined || this.props.NumberOfNews <= 0 ? 10 : this.props.NumberOfNews)
//     .map((data, index) => {
//       CarouselElement.push({
//         imageSrc: data.BannerImageUrl.Url,
//         title: data.Title,
//         description: data.Description,
//         url: 'https://artshouselimited.sharepoint.com' + data.File["ServerRelativeUrl"],
//         showDetailsOnHover: false,
//         imageFit: ImageFit.cover,
//         imgStyle: {
//           height: '338px !important',
//           width: '100% !important',
//           zIndex: 0,
//         }
//       });
//       imageURL.push(
//         data.BannerImageUrl.Url,
//       );
//     });
//   }


    
//   console.log(`Collection ${mergedCollection}`);
//   return (
    
//     <div className={styles.pnpImageCarousel11}>
//     <div >
//       <p className={`${styles.NewsHeader11}`}>Announcements</p>
//       <a style={{ color: "rgb(3, 120, 124)" }} href="https://artshouselimited.sharepoint.com/sites/AHLMainHub/SitePages/News.aspx" className={styles['see-all']}>See All</a>
//     </div>
//     <Carousel 
//       ///add interval
//       interval={this.props.NumberOfIntervelInSeconds * 1000}
//       buttonsLocation={CarouselButtonsLocation.center}
//       buttonsDisplay={CarouselButtonsDisplay.buttonsOnly}
//       // indicatorsDisplay={CarouselIndicatorsDisplay.overlap} 
//       indicatorShape={CarouselIndicatorShape.circle}
  
//        contentContainerStyles={styles.carouselContent11}
//        containerButtonsStyles={styles.carouselButtonsContainer11}
  
//       isInfinite={this.props.StopScroll===null?true:this.props.StopScroll}
  
//       element={ CarouselElement }
//       onMoveNextClicked={(index: number) => { console.log(`Next button clicked: ${index}`); }}
//       onMovePrevClicked={(index: number) => { console.log(`Prev button clicked: ${index}`); }}
//     />
//     <br></br>
//     <img src={imageURL[0]} alt="image" className={`${styles.imageURL}`}/>
   
//     <br></br>
//     <img src={imageURL[0]} alt="image" />
   
//   </div>
  

  
//   );
// }


// }
import * as React from 'react';
import styles from './CarouselWithSharepoint.module.scss';
import { ICarouselWithSharepointProps } from './ICarouselWithSharepointProps';
import "bootstrap/dist/css/bootstrap.css";
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { ImageFit } from 'office-ui-fabric-react';
import { CarouselIndicatorShape } from '@pnp/spfx-controls-react/lib/controls/carousel';
import { ICarouselImageProps } from '@pnp/spfx-controls-react/lib/controls/carousel';
import { Carousel } from "@pnp/spfx-controls-react/lib/Carousel";
import { CarouselButtonsDisplay } from "@pnp/spfx-controls-react/lib/Carousel";
import { CarouselButtonsLocation } from "@pnp/spfx-controls-react/lib/Carousel";
import { ImageHelper, IImageHelperRequest } from "@microsoft/sp-image-helper";

//API 1
export interface ISliderCarouselListItem {
  Title: string;
  Description: string;
  BannerImageUrl: { Url: string };
  PromotedState: number;
  isLive : string;
  isExpired : boolean;
  File: {ServerRelativeUrl:string}; 
  Created:string;
}

export interface ISliderCarouselDemoState {
  value: ISliderCarouselListItem[];
}

//API 2
export interface ISliderCarouselListItemCollection2 {
  Title: string;
  Description: string;
  BannerImageUrl: { Url: string };
  PromotedState: number;
  isLive : string;
  isExpired : boolean;
  File: {ServerRelativeUrl:string}; 
  Created:string;
}

export interface ISliderCarouselDemoStateCollection2 {
  valueCollection2: ISliderCarouselListItemCollection2[];
}

//API 3
export interface ISliderCarouselListItemCollection3 {
  Title: string;
  Description: string;
  BannerImageUrl: { Url: string };
  PromotedState: number;
  isLive : string;
  isExpired : boolean;
  File: {ServerRelativeUrl:string}; 
  Created:string;
}

export interface ISliderCarouselDemoStateCollection3 {
  valueCollection3: ISliderCarouselListItemCollection3[];
}

export default class CarouselWithSharepoint extends React.Component<
  ICarouselWithSharepointProps,
  { value: ISliderCarouselListItem[],valueCollection2:ISliderCarouselListItemCollection2[],valueCollection3:ISliderCarouselListItemCollection3[]}
> { 
  constructor(props: ICarouselWithSharepointProps) {
    super(props);
    this.state = {
      value: [],
      valueCollection2:[],
      valueCollection3:[]
    };
  }
  private getCarouselListContent = () => {
    try {
      const requestUrl1 = `https://artshouselimited.sharepoint.com/sites/AHLMainHub/_api/web/lists/getbytitle('Site%20Pages')/items?$Select*,File/*&$expand=File`;
      const requestUrl2 = `https://artshouselimited.sharepoint.com/sites/AHLHRHub/_api/web/lists/getbytitle('Site%20Pages')/items?$Select*,File/*&$expand=File`;
      const requestUrl3 = `https://artshouselimited.sharepoint.com/sites/AHLITHub/_api/web/lists/getbytitle('Site%20Pages')/items?$Select*,File/*&$expand=File`;
      Promise.all([
        this.props.spHttpClient.get(requestUrl1, SPHttpClient.configurations.v1),
        this.props.spHttpClient.get(requestUrl2, SPHttpClient.configurations.v1),
        this.props.spHttpClient.get(requestUrl3, SPHttpClient.configurations.v1)
      ])
        .then(([response1, response2, response3]: [SPHttpClientResponse, SPHttpClientResponse, SPHttpClientResponse]) => {
          if (response1.ok && response2.ok && response3.ok) {
            debugger
            return Promise.all([response1.json(), response2.json(), response3.json()]);
          }
        })
        .then(([data1, data2, data3]: [{ value: any }, { value: any }, { value: any }]) => {
          if (data1 != null && data2 != null && data3 != null) {
            this.setState({
              value: data1.value,
              valueCollection2: data2.value,
              valueCollection3: data3.value
            });
          }
        });
    } catch (error) {
      console.log('error in service ', error);
    }
  };
  

  componentDidMount = () => {
    this.getCarouselListContent();
  };

//success merge 3 collection
public render(): React.ReactElement<ICarouselWithSharepointProps> {
  let collection = this.state.value;
  let collection2 = this.state.valueCollection2;
  let collection3 = this.state.valueCollection3;

  let imageURL: string[] =[];
  
  // Merge the three collections
  let mergedCollection = collection.concat(collection2, collection3);
  let CarouselElement: ICarouselImageProps[] = [];
  if (mergedCollection.length > 0) {
    // Sort the list items by the Created field in desending order
    mergedCollection.sort((a, b) => new Date(b.Created).getTime() - new Date(a.Created).getTime());

    mergedCollection
        .filter(item => item.PromotedState === 2 && item.isLive === "Yes")
        .slice(0, this.props.NumberOfNews === undefined || this.props.NumberOfNews <= 0 ? 10 : this.props.NumberOfNews)
        .map((data, index) => {
          
          const imageHelperRequest: IImageHelperRequest = {
            sourceUrl: data.File.ServerRelativeUrl,
            // height: 500,
            width: 1000
        };
        let resizedImage = ImageHelper.convertToImageUrl(imageHelperRequest);

            CarouselElement.push({
                imageSrc: resizedImage,
                title: data.Title,
                description: data.Description,
                url: 'https://artshouselimited.sharepoint.com' + data.File.ServerRelativeUrl,
                showDetailsOnHover: false,
                imageFit: ImageFit.cover,
                imgStyle: {
                    height: '338px !important',
                    width: '100% !important',
                    zIndex: 0,
                }
            });
            imageURL.push(data.BannerImageUrl.Url);
            resizedImage = '';
        });
}

    
  console.log(`Collection ${mergedCollection}`);
  return (
    
    <div className={styles.pnpImageCarousel11}>
    <div >
      <p className={`${styles.NewsHeader11}`}>Announcements</p>
      <a style={{ color: "rgb(3, 120, 124)" }} href="https://artshouselimited.sharepoint.com/sites/AHLMainHub/SitePages/News.aspx" className={styles['see-all']}>See All</a>
    </div>
    <Carousel 
      ///add interval
      interval={this.props.StopScroll ===true? null: this.props.NumberOfIntervelInSeconds=== null ? 5 * 1000: this.props.NumberOfIntervelInSeconds * 1000}
      buttonsLocation={CarouselButtonsLocation.center}
      buttonsDisplay={CarouselButtonsDisplay.buttonsOnly}
      // indicatorsDisplay={CarouselIndicatorsDisplay.overlap} 
      indicatorShape={CarouselIndicatorShape.circle}
  
       contentContainerStyles={styles.carouselContent11}
       containerButtonsStyles={styles.carouselButtonsContainer11}
  
      isInfinite={true}
  
      element={ CarouselElement }
      onMoveNextClicked={(index: number) => { console.log(`Next button clicked: ${index}`); }}
      onMovePrevClicked={(index: number) => { console.log(`Prev button clicked: ${index}`); }}
    />
 
  </div>
  );
}
}