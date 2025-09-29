require "test_helper"

class FileUploadsControllerTest < ActionDispatch::IntegrationTest
  test "should get create" do
    get file_uploads_create_url
    assert_response :success
  end
end
