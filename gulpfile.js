const { src, dest } = require('gulp');

function buildIcons() {
  return src('nodes/OnlyOffice/*.svg')
    .pipe(dest('dist/icons/OnlyOffice'))
    .pipe(dest('dist/nodes/OnlyOffice'));
}

function buildStructure() {
  // Files are already compiled to the correct structure
  return Promise.resolve();
}

exports['build:icons'] = buildIcons;
exports['build:structure'] = buildStructure;
