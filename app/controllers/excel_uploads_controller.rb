class ExcelUploadsController < ApplicationController
  require 'roo'
  require 'securerandom'

  def upload
    header = request.headers['Authorization']
    token = header.split(' ').last if header
    if token.blank?
      render json: { error: 'Authorization token missing' }, status: :unauthorized
      return
    end

    decoded = JsonWebToken.decode(token)
    if decoded.nil?
      render json: { error: 'Invalid or expired token' }, status: :unauthorized
      return
    end

    google_token = decoded["google_token"]
    uploaded_file = params[:file]
    return render json: { error: "No file uploaded" }, status: :bad_request unless uploaded_file

    # Save the uploaded file temporarily
    temp_file_path = Rails.root.join("tmp", "#{SecureRandom.uuid}.xlsx")
    File.open(temp_file_path, "wb") { |f| f.write(uploaded_file.read) }

    # Generate a safe job ID
    job_id = SecureRandom.uuid

    # Enqueue the background job asynchronously
    PowerPointGeneratorJob.perform_later(temp_file_path.to_s, job_id, google_token)

    # Respond with download endpoint
    render json: {
      message: "Upload received. Your PowerPoint is being generated.",
      download_url: "/excel_uploads/download/#{job_id}"
    }, status: :accepted
  end

  def download
    raw_id = params[:id].to_s
    safe_id = raw_id.gsub(/[^0-9a-z\-]/i, '')
    pptx_path = Rails.root.join("tmp", "generated_slide_#{safe_id}.pptx")

    if File.exist?(pptx_path)
      send_file pptx_path,
                filename: "presentation_#{safe_id}.pptx",
                type: "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    else
      render json: { message: "File not ready yet or does not exist." }, status: :not_found
    end
  end
end