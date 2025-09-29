class AuthController < ApplicationController
  def login
    google_token = params[:google_token]
    # TODO: verify google_token with google api if google_token is valid genearate token and response to client
    # TODO: if google_token is not valid return unauthorised
    # return render json: {
    #   message: "Unauthorised",
    # }, status: :authorised
    token = JsonWebToken.encode(google_token: google_token)

    # Respond with download endpoint
    render json: {
      message: "Token is generated successfully.",
      token: token
    }, status: :ok
  end
end