require "test_helper"

class ExcelUploadsControllerTest < ActionDispatch::IntegrationTest
  test "should get create" do
    get excel_uploads_create_url
    assert_response :success
  end
end
