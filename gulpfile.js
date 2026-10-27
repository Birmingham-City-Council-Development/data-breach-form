'use strict';

const fs = require('fs');
const path = require('path');
const webpack = require('webpack');
const build = require('@microsoft/sp-build-web');

// ── Deployment settings from .env ────────────────────────────────────────────
// Copy .env.example to .env (gitignored) and fill in your organisation's values.
// Real environment variables win over .env, so CI can set them directly.
// Values are baked into the bundle at build time — re-run the build after
// changing .env. Don't put API keys here: they'd end up inside the .sppkg.
// Keys stay in each web part's property pane.
const APP_ENV_KEYS = [
  'ORG_NAME',
  'TEAM_NAME',
  'REVIEW_TEAM_NAME',
  'ADMIN_NOTIFICATION_EMAIL',
  'SP_LIST_NAME',
  'SP_ADMIN_GROUPS',
  'SP_REVIEWER_GROUPS',
  'LIVE_SERVICE_URL',
  'MAX_ATTACHMENT_MB',
  'ALLOWED_ATTACHMENT_EXTENSIONS',
  'SENDMAIL_URL',
  'APPEND_COMM_LOG_URL',
  'UPDATE_PANEL_RECOMMENDATIONS_URL',
  'UPLOAD_ATTACHMENT_URL',
];

function loadEnvFile(file) {
  if (!fs.existsSync(file)) return false;
  fs.readFileSync(file, 'utf8').split(/\r?\n/).forEach(function (line) {
    const match = line.match(/^\s*([A-Z0-9_]+)\s*=\s*(.*?)\s*$/);
    if (!match) return;
    let value = match[2];
    if (/^(['"]).*\1$/.test(value)) value = value.slice(1, -1);
    if (process.env[match[1]] === undefined) process.env[match[1]] = value;
  });
  return true;
}

if (!loadEnvFile(path.join(__dirname, '.env'))) {
  console.warn('[env] No .env file found — copy .env.example to .env and fill it in. Building with defaults.');
}

// config/serve.json's {tenantDomain} placeholder is filled from this.
if (!process.env.SPFX_SERVE_TENANT_DOMAIN && process.env.SP_SITE_URL) {
  process.env.SPFX_SERVE_TENANT_DOMAIN = process.env.SP_SITE_URL.replace(/^https?:\/\//, '').replace(/\/+$/, '');
}

build.configureWebpack.mergeConfig({
  additionalConfiguration: function (config) {
    const definitions = {};
    APP_ENV_KEYS.forEach(function (key) {
      definitions['process.env.' + key] = JSON.stringify(process.env[key] || '');
    });
    config.plugins.push(new webpack.DefinePlugin(definitions));

    // GDS Transport is licensed for GOV.UK services only — strip it from
    // GOV.UK Frontend's CSS so pages fall back to Arial. See the loader.
    config.module.rules.push({
      test: /govuk-frontend\.min\.css$/,
      enforce: 'pre',
      use: [path.join(__dirname, 'tools', 'strip-gds-transport-loader.js')],
    });
    return config;
  }
});

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

build.initialize(require('gulp'));
