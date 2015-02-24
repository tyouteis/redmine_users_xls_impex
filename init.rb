require 'redmine'
require 'dispatcher' unless Rails::VERSION::MAJOR >= 3
require 'uiex_asset_helpers'

if Rails::VERSION::MAJOR >= 3
  ActionDispatch::Callbacks.to_prepare do
    Mime::Type.register('application/vnd.ms-excel', :xls, %w(application/vnd.ms-excel)) unless defined?(Mime::XLS)
  end
else
  Dispatcher.to_prepare UIEX_AssetHelpers::PLUGIN_NAME do
    Mime::Type.register('application/vnd.ms-excel', :xls, %w(application/vnd.ms-excel)) unless defined?(Mime::XLS)
  end
end

Redmine::Plugin.register UIEX_AssetHelpers::PLUGIN_NAME do
  name 'Users XLS import/export'
  author 'Vitaly Klimov'
  author_url 'mailto:vitaly.klimov@snowbirdgames.com'
  description 'This plugin requires spreadsheet gem.'
  version '0.1.1'

  requires_redmine :version_or_higher => '1.3.0'
end
