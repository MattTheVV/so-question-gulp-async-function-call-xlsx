var gulp = require('gulp')
var paths = require('../paths');

var plumber = require('gulp-plumber');
var replaceExt = require('replace-ext');
var PluginError = require('plugin-error')
var gulpFunction = require('gulp-function').gulpFunction // default ES6 export
var through = require("through2")
var convertExcel = require('excel-as-json').processFile;
var assign = Object.assign || require('object.assign');
var notify = require('gulp-notify');
var gulpIgnore = require('gulp-ignore');
var rename = require('gulp-rename');

gulp.task('excel-to-jsoncsv', function() {
  return gulp.src(paths.sourcefiles_glob)
    .pipe(gulpIgnore.exclude("*\~*")) // Ignore temporary files  by Excel while xlsx is  open
    .pipe(gulpIgnore.exclude("*\$*")) // Ignore temporary files  by Excel while xlsx is  open
    .pipe(plumber({errorHandler: notify.onError('gulp-excel-to-jsoncsv error: <%= error.message %>')}))
    .pipe(GulpConvertExcelToJson()) // This is where the magic should happen
    .pipe(rename({ extname: '.csv' })) // Saving as .csv for SharePoint (does not allow .json files to be saved)
    .pipe(gulp.dest(paths.targetfiles_path))
});

function GulpConvertExcelToJson() {
  return through.obj(function(chunk, enc, callback) {
    var self = this
    if (chunk.isNull() || chunk.isDirectory()) {
      callback() // Do not process directories
      // return
    }

    if (chunk.isStream()) {
      callback() // Do not process streams
      // return
    }

    if (chunk.isBuffer()) {
      convertExcel(chunk.path, null, null, // Converts file found at `chunk.path` and returns (err, `data`) its callback.
        function(err, data) {
          if (err) {
            callback(new PluginError("Excel-as-json", err))
          }
          chunk.contents = new Buffer(JSON.stringify(data))
          self.push(chunk)
          callback()
          // return
        })
    } else {
      callback()
    }
  })
}
