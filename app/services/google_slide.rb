require 'google/apis/slides_v1'
require 'google/apis/drive_v3'
require 'googleauth'

class GoogleSlide
  SLIDES = Google::Apis::SlidesV1
  DRIVE  = Google::Apis::DriveV3

  def initialize(oauth_token)
    @slides_service = SLIDES::SlidesService.new
    @drive_service  = DRIVE::DriveService.new

    # Authenticate both services with OAuth token
    @slides_service.authorization = build_authorization(oauth_token)
    @drive_service.authorization  = build_authorization(oauth_token)
  end

  # Create a new presentation with a given title
#   def create_presentation(title)
#     presentation = SLIDES::Presentation.new(title: title)
#     @slides_service.create_presentation(presentation)
#   end

  # Get a presentation by ID
  def get_presentation(presentation_id)
    @slides_service.get_presentation(presentation_id)
  end

  # Add a blank slide to the presentation
  def add_slide(presentation_id)
    request = SLIDES::Request.new(
      create_slide: SLIDES::CreateSlideRequest.new
    )
    batch_request = SLIDES::BatchUpdatePresentationRequest.new(requests: [request])
    @slides_service.batch_update_presentation(presentation_id, batch_request)
  end

  private

  # Build authorization from OAuth token
  def build_authorization(oauth_token)
    Signet::OAuth2::Client.new(
      access_token: oauth_token
    )
  end
end
