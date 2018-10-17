'use strict';

const gulp = require('gulp');
const path = require('path');
const build = require('@microsoft/sp-build-web');
const webpack = require('webpack');
const bundleAnalyzer = require('webpack-bundle-analyzer');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);
build.addSuppression(/Warning\s-\s[\w]+ApplicationCustomizer\: Admins can make this solution available to all sites in the organization/);
build.addSuppression(/Warning\s-\s[\w]+FieldCustomizer\: Admins can make this solution available to all sites in the organization/);
build.addSuppression(/Warning\s-\s[\w]+CommandSet\: Admins can make this solution available to all sites in the organization/);

/**
 * Builds sourcemaps, so we can debug production solutions
 * @see https://blog.mastykarz.nl/debug-production-version-sharepoint-framework-solution/?utm_source=collab365&utm_medium=collab365today&utm_campaign=daily_digest 
build.configureWebpack.mergeConfig({
  additionalConfiguration: (generatedConfiguration) => {
    generatedConfiguration.devtool = 'source-map';

    for (var i = 0; i < generatedConfiguration.plugins.length; i++) {
      const plugin = generatedConfiguration.plugins[i];
      if (plugin instanceof webpack.optimize.UglifyJsPlugin) {
        plugin.options.sourceMap = true;
        break;
      }
    }

    return generatedConfiguration;
  }
});*/

/**
 * Analyses bundle contents.
 * Useful for moving modules to externals. 
 *
build.configureWebpack.mergeConfig({
  additionalConfiguration: (generatedConfiguration) => {
    const lastDirName = path.basename(__dirname);
    const dropPath = path.join(__dirname, 'temp', 'stats');
    generatedConfiguration.plugins.push(new bundleAnalyzer.BundleAnalyzerPlugin({
      openAnalyzer: false,
      analyzerMode: 'static',
      reportFilename: path.join(dropPath, `${lastDirName}.stats.html`),
      generateStatsFile: true,
      statsFilename: path.join(dropPath, `${lastDirName}.stats.json`),
      logLevel: 'error'
    }));

    return generatedConfiguration;
  }
});*/

require('./gulpfile-custom-tasks');

build.initialize(gulp);