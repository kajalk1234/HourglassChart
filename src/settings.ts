/*
 *  Power BI Visualizations
 *
 *  Copyright (c) Microsoft Corporation
 *  All rights reserved.
 *  MIT License
 *
 *  Permission is hereby granted, free of charge, to any person obtaining a copy
 *  of this software and associated documentation files (the "Software"), to deal
 *  in the Software without restriction, including without limitation the rights
 *  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 *  copies of the Software, and to permit persons to whom the Software is
 *  furnished to do so, subject to the following conditions:
 *
 *  The above copyright notice and this permission notice shall be included in
 *  all copies or substantial portions of the Software.
 *
 *  THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 *  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 *  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 *  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 *  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 *  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 *  THE SOFTWARE.
 */

module powerbi.extensibility.visual {
  'use strict';
  import DataViewObjectsParser = powerbi.extensibility.utils.dataview.DataViewObjectsParser;
  import LegendPosition = powerbi.extensibility.utils.chart.legend.LegendPosition;

  export let chartProperties: {
    rotationSettings: {
      show: DataViewObjectPropertyIdentifier;
    };
    animationSettings: {
      show: DataViewObjectPropertyIdentifier;
    };
    orientationSettings: {
      show: DataViewObjectPropertyIdentifier;
      showFlip: DataViewObjectPropertyIdentifier;
    };
    legend: {
      fontSize: DataViewObjectPropertyIdentifier;
      fontFamily: DataViewObjectPropertyIdentifier;
      labelColor: DataViewObjectPropertyIdentifier;
      legendName: DataViewObjectPropertyIdentifier;
      position: DataViewObjectPropertyIdentifier;
      show: DataViewObjectPropertyIdentifier;
      showTitle: DataViewObjectPropertyIdentifier;
      sizeLegendColor: DataViewObjectPropertyIdentifier;
      displayUnits: DataViewObjectPropertyIdentifier;
      decimalPlaces: DataViewObjectPropertyIdentifier;
    };
    colorSettings: {
      fill: DataViewObjectPropertyIdentifier;
    };
    categoryLabelSettings: {
      show: DataViewObjectPropertyIdentifier;
      fontColor: DataViewObjectPropertyIdentifier;
      fontFamily: DataViewObjectPropertyIdentifier;
      fontSize: DataViewObjectPropertyIdentifier;
      textwrap: DataViewObjectPropertyIdentifier;
    };
    detailLabelSettings: {
      show: DataViewObjectPropertyIdentifier;
      fontSize: DataViewObjectPropertyIdentifier;
      fontFamily: DataViewObjectPropertyIdentifier;
      color: DataViewObjectPropertyIdentifier;
      labelDisplayUnits: DataViewObjectPropertyIdentifier;
      labelPrecision: DataViewObjectPropertyIdentifier;
    };
    separatorSettings: {
      show: DataViewObjectPropertyIdentifier;
      color: DataViewObjectPropertyIdentifier;
      strokeWidth: DataViewObjectPropertyIdentifier;
      lineType: DataViewObjectPropertyIdentifier;
    };
    conversionSettings: {
      show: DataViewObjectPropertyIdentifier;
      label: DataViewObjectPropertyIdentifier;
      labelfontColor: DataViewObjectPropertyIdentifier;
      fontSize: DataViewObjectPropertyIdentifier;
      fontColor: DataViewObjectPropertyIdentifier;
      labelPrecision: DataViewObjectPropertyIdentifier;
      transparency: DataViewObjectPropertyIdentifier;
      displayValue: DataViewObjectPropertyIdentifier;
    };
    gradientSettings: {
      show: DataViewObjectPropertyIdentifier;
    }
  } = {
    rotationSettings: {
      show: <DataViewObjectPropertyIdentifier>{ objectName: 'rotationSettings', propertyName: 'show' }
    },
    animationSettings: {
      show: <DataViewObjectPropertyIdentifier>{ objectName: 'animationSettings', propertyName: 'show' }
    },
    orientationSettings: {
      show: <DataViewObjectPropertyIdentifier>{ objectName: 'orientationSettings', propertyName: 'show' },
      showFlip: <DataViewObjectPropertyIdentifier>{ objectName: 'orientationSettings', propertyName: 'showFlip' }
    },
    legend: {
      fontSize: <DataViewObjectPropertyIdentifier>{ objectName: 'legend', propertyName: 'fontSize' },
      fontFamily: <DataViewObjectPropertyIdentifier>{ objectName: 'legend', propertyName: 'fontFamily' },
      labelColor: <DataViewObjectPropertyIdentifier>{ objectName: 'legend', propertyName: 'labelColor' },
      legendName: <DataViewObjectPropertyIdentifier>{ objectName: 'legend', propertyName: 'titleText' },
      position: <DataViewObjectPropertyIdentifier>{ objectName: 'legend', propertyName: 'position' },
      show: <DataViewObjectPropertyIdentifier>{ objectName: 'legend', propertyName: 'show' },
      showTitle: <DataViewObjectPropertyIdentifier>{ objectName: 'legend', propertyName: 'showTitle' },
      sizeLegendColor: <DataViewObjectPropertyIdentifier>{ objectName: 'legend', propertyName: 'sizeLegendColor' },
      displayUnits: <DataViewObjectPropertyIdentifier>{ objectName: 'legend', propertyName: 'displayUnits' },
      decimalPlaces: <DataViewObjectPropertyIdentifier>{ objectName: 'legend', propertyName: 'decimalPlaces' }
    },
    colorSettings: {
      fill: <DataViewObjectPropertyIdentifier>{ objectName: 'colorSettings', propertyName: 'fill' }
    },
    categoryLabelSettings: {
      show: <DataViewObjectPropertyIdentifier>{ objectName: 'categoryLabelSettings', propertyName: 'show' },
      fontColor: <DataViewObjectPropertyIdentifier>{ objectName: 'categoryLabelSettings', propertyName: 'fontColor' },
      fontFamily: <DataViewObjectPropertyIdentifier>{ objectName: 'categoryLabelSettings', propertyName: 'fontFamily' },
      fontSize: <DataViewObjectPropertyIdentifier>{ objectName: 'categoryLabelSettings', propertyName: 'fontSize' },
      textwrap: <DataViewObjectPropertyIdentifier>{ objectName: 'categoryLabelSettings', propertyName: 'textwrap' }
    },
    detailLabelSettings: {
      show: <DataViewObjectPropertyIdentifier>{ objectName: 'detailLabelSettings', propertyName: 'show' },
      fontSize: <DataViewObjectPropertyIdentifier>{ objectName: 'detailLabelSettings', propertyName: 'fontSize' },
      fontFamily: <DataViewObjectPropertyIdentifier>{ objectName: 'detailLabelSettings', propertyName: 'fontFamily' },
      color: <DataViewObjectPropertyIdentifier>{ objectName: 'detailLabelSettings', propertyName: 'color' },
      labelDisplayUnits: <DataViewObjectPropertyIdentifier>{ objectName: 'detailLabelSettings', propertyName: 'labelDisplayUnits' },
      labelPrecision: <DataViewObjectPropertyIdentifier>{ objectName: 'detailLabelSettings', propertyName: 'labelPrecision' }
    },
    separatorSettings: {
      show: <DataViewObjectPropertyIdentifier>{ objectName: 'separatorSettings', propertyName: 'show' },
      color: <DataViewObjectPropertyIdentifier>{ objectName: 'separatorSettings', propertyName: 'color' },
      strokeWidth: <DataViewObjectPropertyIdentifier>{ objectName: 'separatorSettings', propertyName: 'strokeWidth' },
      lineType: <DataViewObjectPropertyIdentifier>{ objectName: 'separatorSettings', propertyName: 'lineType' }
    },
    conversionSettings: {
      show: <DataViewObjectPropertyIdentifier>{ objectName: 'conversionSettings', propertyName: 'show' },
      label: <DataViewObjectPropertyIdentifier>{ objectName: 'conversionSettings', propertyName: 'label' },
      labelfontColor: <DataViewObjectPropertyIdentifier>{ objectName: 'conversionSettings', propertyName: 'labelfontColor' },
      fontSize: <DataViewObjectPropertyIdentifier>{ objectName: 'conversionSettings', propertyName: 'fontSize' },
      fontColor: <DataViewObjectPropertyIdentifier>{ objectName: 'conversionSettings', propertyName: 'fontColor' },
      labelPrecision: <DataViewObjectPropertyIdentifier>{ objectName: 'conversionSettings', propertyName: 'labelPrecision' },
      transparency: <DataViewObjectPropertyIdentifier>{ objectName: 'conversionSettings', propertyName: 'transparency' },
      displayValue: <DataViewObjectPropertyIdentifier>{ objectName: 'conversionSettings', propertyName: 'displayValue' }
    },
    gradientSettings: {
      show: <DataViewObjectPropertyIdentifier>{ objectName: 'gradientSettings', propertyName: 'show' }
    }
  };

  export function getDefaultColorSettings(): IColorSettings {
    return {
      color: ''
    };
  }

  export function getDefaultCategoryLabelSettings(): ICategoryLabelSettings {
    return {
      show: true,
      fontColor: '#000',
      fontFamily: 'Segoe UI',
      fontSize: 12
    };
  }
  export function getDefaultSeparatorSettings(): ISeparatorSettings {
    return {
      show: false,
      color: '#fff',
      strokeWidth: 4,
      lineType: 'dashed'
    };
  }
  export function getDefaultConversionSettings(): IConversionSettings {
    return {
      show: true,
      label: 'Conversion %',
      labelfontColor: '#000',
      fontSize: 8,
      fontColor: '#000',
      labelPrecision: 0,
      transparency: 0,
      displayValue: 'percent'
    };
  }
  export function getDefaultDetailLabelSettings(): IDetailLabelSettings {
    return {
      show: true,
      fontSize: 12,
      fontFamily: 'Segoe UI',
      color: 'black',
      labelDisplayUnits: 0,
      labelPrecision: 0
    };
  }

  export function getDefaultRotationSettings(): IRotationSettings {
    return {
      show: false
    };
  }

  export function getDefaultGradientSettings(): IGradientSettings {
    return {
      show: false
    };
  }

  export function getDefaultAnimationSettings(): IAnimationSettings {
    return {
      show: false
    };
  }
  export function getDefaultOrientationSettings(): IOrientationSettings {
    return {
      show: false,
      showFlip: false
    };
  }
  export function getDefaultLegendSettings(): ILegendConfig {
    return {
      fontSize: 8,
      fontFamily: 'Segoe UI',
      labelColor: 'rgb(102, 105, 110)',
      legendName: 'Title',
      show: true,
      showTitle: true,
      sizeLegendColor: 'rgba(82, 82, 82, 1)',
      displayUnits: 0,
      decimalPlaces: 0
    };
  }
  export module DataViewObjects {
    /** Gets the value of the given object/property pair. */
    // tslint:disable-next-line:no-shadowed-variable
    export function getValue<T>(objects: DataViewObjects, propertyId: DataViewObjectPropertyIdentifier, defaultValue?: T): T {
      if (!objects) { return defaultValue; }
      const objectOrMap: DataViewObject = objects[propertyId.objectName];

      const object: DataViewObject = <DataViewObject>objectOrMap;

      return DataViewObject.getValue(object, propertyId.propertyName, defaultValue);
    }

    export function getFillColor(
      objects: DataViewObjects,
      propertyId: DataViewObjectPropertyIdentifier,
      defaultColor?: string): string {
      const value: Fill = getValue(objects, propertyId);
      if (!value || !value.solid) { return defaultColor; }

      return value.solid.color;
    }
  }

  export function getColorSettings(dataView: DataView): IColorSettings {
    let objects: DataViewObjects = null;
    const colorSettings: IColorSettings = getDefaultColorSettings();
    if (!dataView.metadata || !dataView.metadata.objects) { return colorSettings; }
    objects = dataView.metadata.objects;
    colorSettings.color = DataViewObjects.getFillColor(objects, chartProperties.colorSettings.fill, colorSettings.color);

    return colorSettings;
  }

  export function getCategoryLabelSettings(dataView: DataView): ICategoryLabelSettings {
    let objects: DataViewObjects = null;
    const categoryLabelSettings: ICategoryLabelSettings = getDefaultCategoryLabelSettings();
    if (!dataView.metadata || !dataView.metadata.objects) { return categoryLabelSettings; }
    objects = dataView.metadata.objects;

    categoryLabelSettings.show = DataViewObjects.getValue(objects, chartProperties.categoryLabelSettings.show,
                                                          categoryLabelSettings.show);
    categoryLabelSettings.fontColor = DataViewObjects.getFillColor(objects, chartProperties.categoryLabelSettings.fontColor,
                                                                   categoryLabelSettings.fontColor);
    categoryLabelSettings.fontFamily = DataViewObjects.getValue(objects, chartProperties.categoryLabelSettings.fontFamily,
                                                                categoryLabelSettings.fontFamily);
    categoryLabelSettings.fontSize = DataViewObjects.getValue(objects, chartProperties.categoryLabelSettings.fontSize,
                                                              categoryLabelSettings.fontSize);

    return categoryLabelSettings;
  }
  export function getSeparatorSettings(dataView: DataView): ISeparatorSettings {
    let objects: DataViewObjects = null;
    const separatorSettings: ISeparatorSettings = getDefaultSeparatorSettings();
    if (!dataView.metadata || !dataView.metadata.objects) { return separatorSettings; }
    objects = dataView.metadata.objects;
    separatorSettings.show = DataViewObjects.getValue(objects, chartProperties.separatorSettings.show, separatorSettings.show);
    separatorSettings.color = DataViewObjects.getFillColor(objects, chartProperties.separatorSettings.color, separatorSettings.color);
    separatorSettings.strokeWidth = DataViewObjects.getValue(objects, chartProperties.separatorSettings.strokeWidth,
                                                             separatorSettings.strokeWidth);
    separatorSettings.lineType = DataViewObjects.getValue(objects, chartProperties.separatorSettings.lineType,
                                                          separatorSettings.lineType);

    return separatorSettings;
  }
  export function getConversionSettings(dataView: DataView): IConversionSettings {
    let objects: DataViewObjects = null;
    const conversionSettings: IConversionSettings = getDefaultConversionSettings();
    if (!dataView.metadata || !dataView.metadata.objects) { return conversionSettings; }
    objects = dataView.metadata.objects;
    conversionSettings.show = DataViewObjects.getValue(objects, chartProperties.conversionSettings.show, conversionSettings.show);
    conversionSettings.label = DataViewObjects.getValue(objects, chartProperties.conversionSettings.label, conversionSettings.label);
    conversionSettings.labelfontColor = DataViewObjects.getValue(objects, chartProperties.conversionSettings.labelfontColor,
                                                                 conversionSettings.labelfontColor);
    conversionSettings.fontSize = DataViewObjects.getValue(objects, chartProperties.conversionSettings.fontSize,
                                                           conversionSettings.fontSize);
    conversionSettings.fontColor = DataViewObjects.getFillColor(objects, chartProperties.conversionSettings.fontColor,
                                                                conversionSettings.fontColor);
    conversionSettings.labelPrecision = DataViewObjects.getValue(objects, chartProperties.conversionSettings.labelPrecision,
                                                                 conversionSettings.labelPrecision);
    conversionSettings.transparency = DataViewObjects.getValue(objects, chartProperties.conversionSettings.transparency,
                                                               conversionSettings.transparency);
    conversionSettings.displayValue = DataViewObjects.getValue(objects, chartProperties.conversionSettings.displayValue,
                                                               conversionSettings.displayValue);

    return conversionSettings;
  }
  export function getDetailLabelSettings(dataView: DataView): IDetailLabelSettings {
    let objects: DataViewObjects = null;
    const detailLabelSettings: IDetailLabelSettings = getDefaultDetailLabelSettings();
    if (!dataView.metadata || !dataView.metadata.objects) { return detailLabelSettings; }
    objects = dataView.metadata.objects;
    detailLabelSettings.fontSize = DataViewObjects.getValue(objects, chartProperties.detailLabelSettings.fontSize,
                                                            detailLabelSettings.fontSize);
    detailLabelSettings.color = DataViewObjects.getFillColor(objects, chartProperties.detailLabelSettings.color, detailLabelSettings.color);
    detailLabelSettings.labelDisplayUnits = DataViewObjects.getValue(objects, chartProperties.detailLabelSettings.labelDisplayUnits,
                                                                     detailLabelSettings.labelDisplayUnits);
    detailLabelSettings.labelPrecision = DataViewObjects.getValue(objects, chartProperties.detailLabelSettings.labelPrecision,
                                                                  detailLabelSettings.labelPrecision);
    detailLabelSettings.show = DataViewObjects.getValue(objects, chartProperties.detailLabelSettings.show, detailLabelSettings.show);

    return detailLabelSettings;
  }

  export function getRotationSettings(dataView: DataView): IRotationSettings {
    let objects: DataViewObjects = null;
    const rotationSettings: IRotationSettings = getDefaultRotationSettings();
    if (!dataView.metadata || !dataView.metadata.objects) { return rotationSettings; }
    objects = dataView.metadata.objects;
    rotationSettings.show = DataViewObjects.getValue(objects, chartProperties.rotationSettings.show, rotationSettings.show);

    return rotationSettings;
  }
  export function getAnimationSettings(dataView: DataView): IAnimationSettings {
    let objects: DataViewObjects = null;
    const animationSettings: IAnimationSettings = getDefaultAnimationSettings();
    if (!dataView.metadata || !dataView.metadata.objects) { return animationSettings; }
    objects = dataView.metadata.objects;
    animationSettings.show = DataViewObjects.getValue(objects, chartProperties.animationSettings.show, animationSettings.show);

    return animationSettings;
  }
  export function getOrientationSettings(dataView: DataView): IOrientationSettings {

    let objects: DataViewObjects = null;
    const orientationSettings: IOrientationSettings = getDefaultOrientationSettings();
    if (!dataView.metadata || !dataView.metadata.objects) { return orientationSettings; }
    objects = dataView.metadata.objects;
    orientationSettings.show = DataViewObjects.getValue(objects, chartProperties.orientationSettings.show, orientationSettings.show);
    orientationSettings.showFlip = DataViewObjects.getValue(objects, chartProperties.orientationSettings.showFlip,
                                                            orientationSettings.showFlip);

    return orientationSettings;
  }

  export function getLegendSettings(dataView: DataView): ILegendConfig {
    let objects: DataViewObjects = null;
    const legendSetting: ILegendConfig = getDefaultLegendSettings();
    if (!dataView.metadata || !dataView.metadata.objects) { return legendSetting; }
    objects = dataView.metadata.objects;
    legendSetting.show = DataViewObjects
      .getValue<boolean>(objects, chartProperties.legend.show, legendSetting.show);
    legendSetting.legendName = DataViewObjects
      .getValue(objects, chartProperties.legend.legendName, legendSetting.legendName);
    legendSetting.showTitle = <boolean>DataViewObjects
      .getValue(objects, chartProperties.legend.showTitle, legendSetting.showTitle);
    legendSetting.fontSize = DataViewObjects
      .getValue(objects, chartProperties.legend.fontSize, legendSetting.fontSize);
    legendSetting.fontFamily = DataViewObjects
      .getValue(objects, chartProperties.legend.fontFamily, legendSetting.fontFamily);
    legendSetting.labelColor = DataViewObjects
      .getFillColor(objects, chartProperties.legend.labelColor, legendSetting.labelColor);
    legendSetting.sizeLegendColor = DataViewObjects
      .getFillColor(objects, chartProperties.legend.sizeLegendColor, legendSetting.sizeLegendColor);
    legendSetting.displayUnits = DataViewObjects
      .getValue(objects, chartProperties.legend.displayUnits, legendSetting.displayUnits);
    legendSetting.decimalPlaces = DataViewObjects
      .getValue(objects, chartProperties.legend.decimalPlaces, legendSetting.decimalPlaces);

    if (legendSetting.decimalPlaces > 4) {
      legendSetting.decimalPlaces = 4;
    } else if (legendSetting.decimalPlaces < 0) {
      legendSetting.decimalPlaces = 0;
    }

    return legendSetting;
  }

  export function getGradientSettings(dataView: DataView): IGradientSettings {
    let objects: DataViewObjects = null;
    const gradientSettings: IGradientSettings = getDefaultGradientSettings();
    if (!dataView.metadata || !dataView.metadata.objects) { return gradientSettings; }
    objects = dataView.metadata.objects;
    gradientSettings.show = DataViewObjects.getValue(objects, chartProperties.gradientSettings.show, gradientSettings.show);

    return gradientSettings;
  }

  export module DataViewObject {
    export function getValue<T>(object: DataViewObject, propertyName: string, defaultValue?: T): T {

      if (!object) { return defaultValue; }
      const propertyValue: T = <T>object[propertyName];
      if (propertyValue === undefined) { return defaultValue; }

      return propertyValue;
    }

  }

  // tslint:disable-next-line:no-any
  export function enumerateColorSetting(visualModel: any,
                                        instance: VisualObjectInstance[], objectName: string): void {

    for (const visualDataPoint of visualModel.dataPoints) {
      instance.push({
        objectName: objectName,
        displayName: visualDataPoint.category.toString(),
        properties: {
          color: visualDataPoint.color
        },
        selector: visualDataPoint.selectionId.getSelector()
      });
    }
  }
  export function enumerateCategoryLabelSetting(categoryLabelSettings: ICategoryLabelSettings,
                                                instance: VisualObjectInstance[], objectName: string): void {
    let props: {} = {};
    props = {
      show: categoryLabelSettings.show,
      fontColor: categoryLabelSettings.fontColor,
      fontFamily: categoryLabelSettings.fontFamily,
      fontSize: categoryLabelSettings.fontSize
    };
    instance.push({
      objectName: objectName,
      properties: props,
      selector: null
    });
  }
  export function enumerateSeparatorSetting(separatorSettings: ISeparatorSettings,
                                            instance: VisualObjectInstance[], objectName: string): void {
    let props: {} = {};
    props = {
      show: separatorSettings.show,
      color: separatorSettings.color,
      strokeWidth: separatorSettings.strokeWidth,
      lineType: separatorSettings.lineType
    };
    instance.push({
      objectName: objectName,
      properties: props,
      selector: null
    });
  }
  export function enumerateConversionSetting(conversionSettings: IConversionSettings,
                                             instance: VisualObjectInstance[], objectName: string): void {
    let props: {} = {};
    props = {
      show: conversionSettings.show,
      label: conversionSettings.label,
      labelfontColor: conversionSettings.labelfontColor,
      fontSize: conversionSettings.fontSize,
      fontColor: conversionSettings.fontColor,
      labelPrecision: conversionSettings.labelPrecision,
      transparency: conversionSettings.transparency,
      displayValue: conversionSettings.displayValue
    };
    instance.push({
      objectName: objectName,
      properties: props,
      selector: null
    });
  }
  export function enumerateDetailLabelSetting(detailLabelSettings: IDetailLabelSettings,
                                              instance: VisualObjectInstance[], objectName: string): void {
    let props: {} = {};
    props = {
      show: detailLabelSettings.show,
      fontSize: detailLabelSettings.fontSize,
      color: detailLabelSettings.color,
      labelDisplayUnits: detailLabelSettings.labelDisplayUnits,
      labelPrecision: detailLabelSettings.labelPrecision
    };
    instance.push({
      objectName: objectName,
      properties: props,
      selector: null
    });
  }

  export function enumerateRotationSetting(rotationSettings: IRotationSettings,
                                           instance: VisualObjectInstance[], objectName: string): void {
    let props: {} = {};
    props = {
      show: rotationSettings.show
    };
    instance.push({
      objectName: objectName,
      properties: props,
      selector: null
    });
  }

  export function enumerateAnimationSetting(animationSettings: IAnimationSettings,
                                            instance: VisualObjectInstance[], objectName: string): void {
    let props: {} = {};
    props = {
      show: animationSettings.show
    };
    instance.push({
      objectName: objectName,
      properties: props,
      selector: null
    });
  }

  export function enumerateOrientationSetting(orientationSettings: IOrientationSettings,
                                              instance: VisualObjectInstance[], objectName: string): void {
    let props: {} = {};
    props = {
      show: orientationSettings.show,
      showFlip: orientationSettings.showFlip
    };
    instance.push({
      objectName: objectName,
      properties: props,
      selector: null
    });
  }

  export function enumerateLegend(legendConfig: ILegendConfig, instance: VisualObjectInstance[]
    ,                             objectName: string): void {

    let props: {} = {};
    props = {
      show: legendConfig.show,
      position: LegendPosition[Visual.legend.getOrientation()],
      fontSize: legendConfig.fontSize,
      fontFamily: legendConfig.fontFamily,
      labelColor: legendConfig.labelColor,
      showTitle: legendConfig.showTitle,
      titleText: legendConfig.legendName
    };
    instance.push({
      objectName: objectName,
      properties: props,
      selector: null
    });
  }

}
