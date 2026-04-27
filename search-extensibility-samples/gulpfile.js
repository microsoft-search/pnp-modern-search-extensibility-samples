'use strict';

const gulp = require('gulp');
const build = require('@microsoft/sp-build-web');
build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);
var getTasks = build.rig.getTasks;
build.rig.getTasks = function() {
    var result = getTasks.call(build.rig);

    result.set('serve', result.get('serve-deprecated'));

    return result;
};

const envCheck = build.subTask('environmentCheck', (gulp, config, done) => {

  build.configureWebpack.mergeConfig({
      additionalConfiguration: (generatedConfiguration) => {

          // Remove the default html rule
          generatedConfiguration.module.rules = generatedConfiguration.module.rules.filter(rule => {
              return rule.test.toString() !== '/\\.html$/';
          });

          generatedConfiguration.module.rules.push({
            // Add html loader without minimize so that we can use it for handlebars templates
            test: /\.html$/,
            loader: 'html-loader',
            options: {
              minimize: false
            }
          });
          return generatedConfiguration;
      }
  });

  done();
});

build.rig.addPreBuildTask(envCheck);

build.initialize(gulp);