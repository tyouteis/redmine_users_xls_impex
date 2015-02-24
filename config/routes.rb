if Rails::VERSION::MAJOR >= 3
  RedmineApp::Application.routes.draw do
    match 'users_xls_impex/:action', :to => 'users_xls_impex'
  end
else
  ActionController::Routing::Routes.draw do |map|
    map.connect 'users_xls_impex/:action', :controller => 'users_xls_impex'
  end
end
