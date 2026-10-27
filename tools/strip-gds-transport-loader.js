'use strict';

// Removes the GDS Transport font from GOV.UK Frontend's precompiled CSS.
//
// GDS Transport is licensed only for services on the GOV.UK domain, so
// anything else must not load it. This produces the same CSS as GOV.UK
// Frontend's own settings for non-GOV.UK services:
//   $govuk-include-default-font-face: false;
//   $govuk-font-family: arial, sans-serif;
// done here instead of compiling GOV.UK Frontend from SCSS, because SPFx's
// bundled Sass version is too old for current GOV.UK Frontend releases.
//
// Fails the build if the font is still referenced afterwards (e.g. a future
// GOV.UK Frontend release formats its CSS differently), so it can't silently
// ship again.
module.exports = function stripGdsTransport(source) {
  const output = source
    .replace(/@font-face\s*\{\s*font-family:\s*["']?GDS Transport["']?;[^}]*\}/g, '')
    .replace(/["']?GDS Transport["']?\s*,\s*arial\s*,\s*sans-serif/g, 'arial,sans-serif');

  if (/GDS Transport/.test(output)) {
    this.emitError(new Error(
      'GDS Transport is still referenced in GOV.UK Frontend CSS after stripping. ' +
      'Update tools/strip-gds-transport-loader.js for this govuk-frontend version.'
    ));
  }
  return output;
};
