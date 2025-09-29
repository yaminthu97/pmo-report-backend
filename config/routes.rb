Rails.application.routes.draw do
  get "file_uploads/create"
  get "excel_uploads/create"
  # post "/excel_uploads", to: "excel_uploads#create"
  post   '/excel_uploads',           to: 'excel_uploads#upload'
  post   '/auth/login',           to: 'auth#login'
  get    '/excel_uploads/download/:id', to: 'excel_uploads#download'
  post "/file_uploads", to: "file_uploads#create"
  # Define your application routes per the DSL in https://guides.rubyonrails.org/routing.html

  # Reveal health status on /up that returns 200 if the app boots with no exceptions, otherwise 500.
  # Can be used by load balancers and uptime monitors to verify that the app is live.
  get "up" => "rails/health#show", as: :rails_health_check

  # Defines the root path route ("/")
  # root "posts#index"
end
