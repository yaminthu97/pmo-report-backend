class FileUploadsController < ApplicationController
  def create
    file = params[:file]
    return render json: { error: "No file uploaded" }, status: :bad_request unless file

    xlsx = Roo::Spreadsheet.open(file.tempfile.path)
    sheet = xlsx.sheet(0)

    headers = sheet.row(2)
    return render json: { error: "Missing headers in row 2" }, status: :unprocessable_entity if headers.blank?

    data_rows = []
    (3..sheet.last_row).each do |i|
    row = sheet.row(i)
    next if row.compact.empty?

    row_hash = Hash[headers.zip(row)]
    data_rows << row_hash
    end

    return render json: { error: "No valid data rows found" }, status: :unprocessable_entity if data_rows.empty?

    deck = Powerpoint::Presentation.new
    deck.add_intro("Project Details", "Imported from Excel")

    # Slide design logic
    data_rows.each_with_index do |project, idx|
    title = project["Project名"] || "Untitled Project"
    body_lines = []

      body_lines << "・ 担当PM: #{project["担当PM"]}"
      body_lines << "・ SV結果: #{project["SV結果の解釈\n【スケジュール差異】"]}"
      body_lines << "・ CV結果: #{project["CV結果の解釈\n【コスト差異】"]}"
      body_lines << "・ CPI: #{project["CPI"]} / SPI: #{project["SPI"]}"

      deck.add_textual_slide("Project #{idx + 1}: #{title}", body_lines)
    end

    pptx_path = Rails.root.join("tmp", "detailed_project_#{Time.now.to_i}.pptx")
    deck.save(pptx_path.to_s)

    send_file pptx_path,
              filename: "project_details.pptx",
              type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
              disposition: "attachment"
    rescue => e
    Rails.logger.error("Error creating PowerPoint: #{e.message}")
    render json: { error: "Failed to create PowerPoint" }, status: :internal_server_error
  end
end
