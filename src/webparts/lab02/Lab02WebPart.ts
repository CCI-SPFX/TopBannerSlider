import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  type IPropertyPaneConfiguration,
  PropertyPaneChoiceGroup,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneToggle,
  IPropertyPaneGroup
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
// import { PropertyFieldPicturePicker } from '@pnp/spfx-property-controls/lib/PropertyFieldPicturePicker';

// import { IReadonlyTheme } from '@microsoft/sp-component-base';

// import * as strings from 'Lab02WebPartStrings';
import Lab02 from "./components/Lab02";
import { ILab02Props } from "./components/ILab02Props";

export interface ILab02WebPartProps {
  heroSlider: boolean;
  slidesList: string;
  timer:number;
  showContent:boolean;

  contentAlign: string;
  verticalAlignment:string;
  headingValue: string;
  headingSize: string;
  headingColor: string;
  showHeading: boolean;

  descValue: string;
  descSize: string;
  descColor: string;
  showDesc: boolean;

  buttonText: string;
  buttonTextSize: string;
  buttonBgColor: string;
  buttonRadius: string;
  buttonTextColor: string;
  buttonLink: string;
  showButton: boolean;

  imageUrl: string;
  bgTheme: string;
  itemHeight:string;
  aspectRatio:string;
}

export default class Lab02WebPart extends BaseClientSideWebPart<ILab02WebPartProps> {
  // private _activeGroup: string;

  // private _isDarkTheme: boolean = false;
  // private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<ILab02Props> = React.createElement(
      Lab02,
      {
        slider: this.properties.heroSlider,
        listName: this.properties.slidesList,
        siteUrl:this.context.pageContext.web.absoluteUrl,
        timer: this.properties.timer,
        showContent: this.properties.showContent,

        contentAlignment: this.properties.contentAlign,
        verticalAlign:this.properties.verticalAlignment,
        heading: this.properties.headingValue,
        headingSize: this.properties.headingSize,
        headingColor: this.properties.headingColor,
        headingShow: this.properties.showHeading,

        description: this.properties.descValue,
        descriptionSize: this.properties.descSize,
        descriptionColor: this.properties.descColor,
        descriptionShow: this.properties.showDesc,

        btnText: this.properties.buttonText,
        btnTextSize: this.properties.buttonTextSize,
        btnBgColor: this.properties.buttonBgColor,
        btnBorderRadius: this.properties.buttonRadius,
        BtnTextColor: this.properties.buttonTextColor,
        btnLink: this.properties.buttonLink,
        btnShow: this.properties.showButton,

        imgUrl: this.properties.imageUrl,
        bgType: this.properties.bgTheme,
        itemHeight:this.properties.itemHeight,
        aspectRatio:this.properties.aspectRatio,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  // protected onInit(): Promise<void> {
  //   return this._getEnvironmentMessage().then(message => {
  //     this._environmentMessage = message;
  //   });
  // }
 
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  // protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
  //   return {
  //     pages: [
  //       {
  //         groups: [
  //           {
  //             groupFields: [
  //               PropertyPaneToggle("heroSlider", {
  //                 label: "Enable Slider",
  //                 offText: "No",
  //                 onText: "Yes",
  //               }),
  //             ],
  //           },
  //           {
  //             groupFields: [
  //               PropertyPaneChoiceGroup("contentAlign", {
  //                 label: "Horizontal Alignment",
  //                 options: [
  //                   {
  //                     key: "Left",
  //                     text: "Left",
  //                     iconProps: { officeFabricIconFontName: "AlignLeft" },
  //                   },
  //                   {
  //                     key: "Center",
  //                     text: "Center",
  //                     iconProps: { officeFabricIconFontName: "AlignCenter" },
  //                   },
  //                   {
  //                     key: "Right",
  //                     text: "Right",
  //                     iconProps: { officeFabricIconFontName: "AlignRight" },
  //                   },
  //                 ],
  //               }),
  //             ],
  //           },
  //           {
  //             groupFields: [
  //               PropertyPaneChoiceGroup("verticalAlignment", {
  //                 label: "Vertical Alignment",
  //                 options: [
  //                   {
  //                     key: "Top",
  //                     text: "Top",
  //                     iconProps: { officeFabricIconFontName: "AlignVerticalTop" },
  //                   },
  //                   {
  //                     key: "Middle",
  //                     text: "Center",
  //                     iconProps: {
  //                       officeFabricIconFontName: "AlignVerticalCenter",
  //                     },
  //                   },
  //                   {
  //                     key: "Bottom",
  //                     text: "Bottom",
  //                     iconProps: { officeFabricIconFontName: "AlignVerticalBottom" },
  //                   },
  //                 ],
  //               }),
  //             ],
  //           },

  //           // ✅ Conditional rendering
  //           ...(this.properties.heroSlider
  //             ? [
  //                 { 
  //                   groupFields: [
  //                     PropertyPaneTextField("slidesList", {
  //                       label: "Slides List",
  //                     }),

  //                   // PropertyPaneTextField("Sectionheight", {label: "Section Height"}),
  //                     PropertyPaneSlider("timer", {
  //                       label: "Slider Timer (in seconds)",
  //                       min: 1,
  //                       max: 15,
  //                       step: 1,
  //                     }),
  //                   ],
  //                 },
  //               ]
  //             : [
  //                 {
  //                   groupName: "Banner Content",
  //                   groupFields: [
  //                     PropertyPaneTextField("headingValue", {
  //                       label: "Heading",
  //                     }),
  //                     PropertyPaneTextField("descValue", {
  //                       label: "Description",
  //                     }),
  //                     PropertyPaneTextField("imageUrl", { 
  //                       label: "Image URL" 
  //                     }),
  //                     PropertyPaneTextField("buttonText", {
  //                       label: "Button Text",
  //                     }),
  //                     PropertyPaneTextField("buttonLink", {
  //                       label: "Button Link",
  //                     }),
  //                   ],
  //                 },
  //               ]), 

  //           {
  //             groupFields: [
  //               PropertyPaneDropdown("bgTheme", {
  //                 label: "Theme",
  //                 options: [
  //                   { key: "None", text: "None" },
  //                   { key: "Light", text: "Light" },
  //                   { key: "Dark", text: "Dark" },
  //                 ],
  //               }),
  //             ],
  //           },
  //           {
  //             groupName: "Heading Properties",
  //             groupFields: [
  //               PropertyPaneToggle("showHeading", {
  //                 label: "Show Heading",
  //                 onText: "ON",
  //                 offText: "OFF",
  //               }),

  //               PropertyPaneTextField("headingSize", { label: "Text Size" }),
  //               PropertyPaneTextField("headingColor", { label: "Text Color" }),
  //             ],
  //           },
  //           {
  //             groupName: "Description Properties",
  //             groupFields: [
  //               PropertyPaneToggle("showDesc", {
  //                 label: "Show Description",
  //                 onText: "ON",
  //                 offText: "OFF",
  //               }),
  //               PropertyPaneTextField("descSize", { label: "Text Size" }),
  //               PropertyPaneTextField("descColor", { label: "Text Color" }),
  //             ],
  //           },

  //           {
  //             groupName: "Button Properties",
  //             groupFields: [
  //               PropertyPaneToggle("showButton", {
  //                 label: "Show Button",
  //                 onText: "ON",
  //                 offText: "OFF",
  //               }),
  //               PropertyPaneTextField("buttonTextSize", { label: "Text Size" }),
  //               PropertyPaneTextField("buttonBgColor", {
  //                 label: "Background Color",
  //               }),
  //               PropertyPaneTextField("buttonTextColor", {
  //                 label: "Text Color",
  //               }),
  //               PropertyPaneTextField("buttonRadius", {
  //                 label: "Border Radius",
  //               }),
  //             ],
  //           },
  //         ],
  //       },
  //     ],
  //   };
  // }




  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
  // Groups that only show when showContent is true
  const contentGroups: IPropertyPaneGroup[] = this.properties.showContent
    ? [
          {
            groupFields: [
              PropertyPaneChoiceGroup("contentAlign", {
                label: "Horizontal Alignment",
                options: [
                  {
                    key: "Left",
                    text: "Left",
                    iconProps: { officeFabricIconFontName: "AlignLeft" },
                  },
                  {
                    key: "Center",
                    text: "Center",
                    iconProps: { officeFabricIconFontName: "AlignCenter" },
                  },
                  {
                    key: "Right",
                    text: "Right",
                    iconProps: { officeFabricIconFontName: "AlignRight" },
                  },
                ],
              }),
            ],
          },
          {
            groupFields: [
              PropertyPaneChoiceGroup("verticalAlignment", {
                label: "Vertical Alignment",
                options: [
                  {
                    key: "Top",
                    text: "Top",
                    iconProps: { officeFabricIconFontName: "AlignVerticalTop" },
                  },
                  {
                    key: "Middle",
                    text: "Center",
                    iconProps: {
                      officeFabricIconFontName: "AlignVerticalCenter",
                    },
                  },
                  {
                    key: "Bottom",
                    text: "Bottom",
                    iconProps: {
                      officeFabricIconFontName: "AlignVerticalBottom",
                    },
                  },
                ],
              }),
            ],
          },

        {
          groupFields: [
            PropertyPaneDropdown("bgTheme", {
              label: "Theme",
              options: [
                { key: "None", text: "None" },
                { key: "Light", text: "Light" },
                { key: "Dark", text: "Dark" },
              ],
            }),
          ],
        },
        {
          groupName: "Heading Properties",
          groupFields: [
            PropertyPaneToggle("showHeading", {
              label: "Show Heading",
              onText: "ON",
              offText: "OFF",
            }),
            PropertyPaneTextField("headingSize", { label: "Text Size" }),
            PropertyPaneTextField("headingColor", { label: "Text Color" }),
          ],
        },
        {
          groupName: "Description Properties",
          groupFields: [
            PropertyPaneToggle("showDesc", {
              label: "Show Description",
              onText: "ON",
              offText: "OFF",
            }),
            PropertyPaneTextField("descSize", { label: "Text Size" }),
            PropertyPaneTextField("descColor", { label: "Text Color" }),
          ],
        },
        {
          groupName: "Button Properties",
          groupFields: [
            PropertyPaneToggle("showButton", {
              label: "Show Button",
              onText: "ON",
              offText: "OFF",
            }),
            PropertyPaneTextField("buttonTextSize", { label: "Text Size" }),
            PropertyPaneTextField("buttonBgColor", {
              label: "Background Color",
            }),
            PropertyPaneTextField("buttonTextColor", {
              label: "Text Color",
            }),
            PropertyPaneTextField("buttonRadius", {
              label: "Border Radius",
            }),
          ],
        },
      ]
    : [];

    const sizeGroup: IPropertyPaneGroup = this.properties.showContent
  ? 
   { 
      groupFields: [
        PropertyPaneTextField("itemHeight", {
          label: "Height",
          description: "Example: 400px, 50vh",
        }),
      ],
    }
    :{ 
      groupFields: [
        PropertyPaneTextField("aspectRatio", {
          label: "Aspect Ratio",
          description: "Example: 22/6, 20/6, 16/9,",
        }),
      ],
    }

  return {
    pages: [
      {
        groups: [
          {
            groupFields: [
              PropertyPaneToggle("heroSlider", {
                label: "Enable Slider",
                offText: "No",
                onText: "Yes",
              }),
            ],
          },
          // Slider vs. banner content (unchanged)
          ...(this.properties.heroSlider
            ? [
                {
                  groupFields: [
                    PropertyPaneTextField("slidesList", {
                      label: "Slides List",
                    }),
                    PropertyPaneSlider("timer", {
                      label: "Slider Timer (in seconds)",
                      min: 1,
                      max: 15,
                      step: 1,
                    }),
                  ],
                },
              ]
            : [
                {
                  groupName: "Banner Content",
                  groupFields: [
                    PropertyPaneTextField("headingValue", {
                      label: "Heading",
                    }),
                    PropertyPaneTextField("descValue", {
                      label: "Description",
                    }),
                    PropertyPaneTextField("imageUrl", {
                      label: "Background image URL",
                    }),
                    PropertyPaneTextField("buttonText", {
                      label: "Button Text",
                    }),
                    PropertyPaneTextField("buttonLink", {
                      label: "Button/Banner Link URL",
                    }),
                  ],
                },
              ]),

          // New: Show Content toggle
          {
            groupFields: [
              PropertyPaneToggle("showContent", {
                label: "Show Content",
                onText: "ON",
                offText: "OFF",
              }),
            ],
          },
          sizeGroup,
          // Theme, Heading, Description, Button groups (only when showContent is true)
          ...contentGroups,
        ],
      },
    ],
  };
}
}

// function PropertyPaneAccordion(arg0: string, arg1: { panels: { title: string; fields: IPropertyPaneField<IPropertyPaneTextFieldProps>[]; }[]; }): import("@microsoft/sp-property-pane").IPropertyPaneGroup | import("@microsoft/sp-property-pane").IPropertyPaneConditionalGroup {
//   throw new Error('Function not implemented.');
// }
