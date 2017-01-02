"use strict";

var gulp = require("gulp");
var del = require("del");
var typescript = require("gulp-typescript");
var merge = require('merge-stream');
var path = require('path')
var glob = require('glob')

gulp.task('default', ['build']);

gulp.task('clean', function (cb) {
    return glob('./**/*.ts', function (err, files) {
		var generatedFiles = files.map(function (file) {
		  return file.replace(/.ts$/, '.js*');
		});

		del(generatedFiles, cb);
	});
});


gulp.task('build', function () {
    var tsProject = typescript.createProject("tsconfig.json");
    return tsProject.src()
        .pipe(tsProject())
        .pipe(gulp.dest(function(file) {
			return file.base;
		}));
});

gulp.task('watch', function () {
    gulp.watch('./src/**/*.ts', ['build']);
});

gulp.task('package:clean', function () {
    return del(['package/*/**']);
});

gulp.task('package', ['package:copy'], function () {	
});

gulp.task('package:copy', ['package:clean'], function () {
    var main = gulp.src([
                './extension-icon*.png', 'LICENSE', 'README.md', 'vss-extension.json', 'package.json', 
                'docs/**/*', 
				'**/*.xsd', '**/*.wsdl',
                './InstallMailApp/**/*.js', './InstallMailApp/package.json', './InstallMailApp/task.json', './InstallMailApp/icon.png',
                './UninstallMailApp/**/*.js', './UninstallMailApp/package.json', './UninstallMailApp/task.json', './UninstallMailApp/icon.png',
				"!**/node_modules/**/*",],
                 {base: "."})
        .pipe(gulp.dest("./package"));

    return merge(main);	
});