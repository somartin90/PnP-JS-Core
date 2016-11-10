//******************************************************************************
//* test.js
//*
//* Defines a custom gulp task for executing the unit tests (with mocha) and
//* also reporting on code coverage (with istanbul).
//******************************************************************************

"use strict";

//******************************************************************************
//* DEPENDENCIES
//******************************************************************************

var gulp = require("gulp"),
    mocha = require("gulp-mocha"),
    istanbul = require("gulp-istanbul"),
    tsc = require("gulp-typescript"),
    yargs = require('yargs').argv;

//******************************************************************************
//* TEST
//******************************************************************************

gulp.task("build-tests", ["clean"], function() {
    var src = global.TSWorkspace.Tests.slice(0);

    return gulp.src(src)
        .pipe(global.tsProject())
        .pipe(gulp.dest(global.TSCompiledOutput.TestRootFolder));
});

gulp.task("test", ["build", "build-tests", "istanbul:hook"], function() {
    let path = './build/tests/{path}.test.js';
    return gulp.src(yargs.single ? path.replace('{path}', yargs.single) : global.TSCompiledOutput.JSTestFiles)
        .pipe(mocha({ ui: 'bdd', reporter: 'dot', timeout: 10000 }))
        .pipe(istanbul.writeReports());
});