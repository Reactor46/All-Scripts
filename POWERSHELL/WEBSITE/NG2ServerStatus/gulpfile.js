var gulp = require('gulp'),
    rimraf = require('rimraf'),
    plugins = require('gulp-load-plugins')({
        lazy: true
    }),
    runSequence = require('run-sequence'),
    autoprefixer = require('autoprefixer'),
    postcss = require('gulp-postcss'),
    autoprefixer = require('autoprefixer'),
    cleancss = require('gulp-clean-css'),
    less = require('gulp-less');

var paths = {
    dist: 'dist'
};

gulp.task('clean:dist', function (done) {
    rimraf(paths.dist, done);
});

gulp.task('dev-server', plugins.shell.task('webpack-dev-server --inline --colors --progress --port 3000'));

gulp.task('build', plugins.shell.task([
    'rimraf dist',
    'webpack --config config/webpack.prod.js --progress --colors --profile --bail',
    'pwd',
    'cp assets/*.* dist/assets/'
]));

gulp.task('serve', function (done) {
    runSequence('clean:dist', 'build-less', 'watch-less', 'dev-server', done);
});

gulp.task('build-less', function () {
    var processors = [
        autoprefixer
    ];
    return gulp.src('src/css/main.less')
        .pipe(less())
        .pipe(postcss(processors))
        .pipe(cleancss({
            compatibility: 'ie8'
        }))
        .pipe(gulp.dest('src/css'));
});

gulp.task('watch-less', function () {
    gulp.watch('src/css/**/*.less', ['build-less']);
    // Other watchers
});
