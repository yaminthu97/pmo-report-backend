require "googleauth"
require "google/apis/drive_v3"
require "google/apis/slides_v1"
require "googleauth/stores/file_token_store"
require "fileutils"
require "open-uri"
require "json"
require "date"

module ExportPowerPointSchedulerModule
  include Google::Apis

  TITLE_COVER_SLIDE_INDEX = 0
  EVM_INDICATOR_SLIDE_INDEX = 1
  OVERALL_SUMMARY1_SLIDE_INDEX = 2
  OVERALL_SUMMARY2_SLIDE_INDEX = 3
  OVERALL_SUMMARY3_SLIDE_INDEX = 4
  COST_SUMMARY_SLIDE_INDEX = 5
  COST_SUMMARY_SLIDE2_INDEX = 6
  EVM_SUMMARY_SLIDE_INDEX = 7
  EVM_MONTHLY_SLIDE_INDEX = 8
  
  ENV["http_proxy"] = nil
  ENV["https_proxy"] = nil
  ENV["HTTP_PROXY"] = nil
  ENV["HTTPS_PROXY"] = nil

  # OOB_URI = "urn:ietf:wg:oauth:2.0:oob"
  LOCALHOST_URI = "http://localhost"
  CREDENTIALS_PATH = Rails.root.join("config", "credentials.json")  # OAuth2 file from Google
  TOKEN_PATH = Rails.root.join("config", "token.yaml")              # Will be created after first run
  SCOPE = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/presentations"
  ]

   def self.log_storage_quota(drive_service)
    # Fetch current user's storage quota details
    about = drive_service.get_about(fields: "storageQuota")
    quota = about.storage_quota

    limit_mb = quota.limit.to_f / (1024 * 1024)
    usage_mb = quota.usage.to_f / (1024 * 1024)
    usage_drive_mb = quota.usage_in_drive.to_f / (1024 * 1024)
    usage_trash_mb = quota.usage_in_drive_trash.to_f / (1024 * 1024)

    Rails.logger.info({
      storage: {
        total_mb: limit_mb.round(2),
        used_mb: usage_mb.round(2),
        used_in_drive_mb: usage_drive_mb.round(2),
        used_in_trash_mb: usage_trash_mb.round(2)
      }
    }.to_json)
  end

def self.export_excel_data(records, google_token)
    grouped = records
    records_all = grouped[:all]
    records_week = grouped[:week]
    records_month_unit = grouped[:month_unit]

    # Set up services with user token
    drive_service = DriveV3::DriveService.new
    slides_service = SlidesV1::SlidesService.new
    signet = Signet::OAuth2::Client.new(access_token: google_token)
    drive_service.authorization = slides_service.authorization = signet

    ExportPowerPointSchedulerModule.log_storage_quota(drive_service)# log
    # Find or create "PMO Reports" folder
    folder_name = "PMO Reports"
    folder_id = nil

    query = "mimeType = 'application/vnd.google-apps.folder' and name = '#{folder_name}' and trashed = false"
    result = drive_service.list_files(q: query, fields: "files(id, name)")

    if result.files.any?
      folder_id = result.files.first.id
    else
      folder_metadata = { name: folder_name, mime_type: "application/vnd.google-apps.folder" }
      folder = drive_service.create_file(folder_metadata, supports_all_drives: true)
      folder_id = folder.id
    end

    # Upload the template inside PMO Reports
    template_name = "TESTPMOReportTemplate"
    template_file_id = nil

    query = "name = '#{template_name}' and '#{folder_id}' in parents and trashed = false"
    result = drive_service.list_files(q: query, fields: "files(id, name)")

    if result.files.any?
      template_file_id = result.files.first.id
    else
      file_path = Rails.root.join("app/assets/files", "TESTPMOReportTemplate.pptx").to_s
      uploaded = drive_service.create_file(
        { name: template_name, parents: [ folder_id ], mime_type: "application/vnd.google-apps.presentation" },
        upload_source: file_path,
        content_type: "application/vnd.openxmlformats-officedocument.presentationml.presentation"
      )
      template_file_id = uploaded.id
    end

    # Find or create "ResultSlides" subfolder inside PMO Reports
    query = "name = 'ResultSlides' and mimeType = 'application/vnd.google-apps.folder' and '#{folder_id}' in parents and trashed = false"
    result = drive_service.list_files(q: query, fields: "files(id, name)")

    if result.files.any?
      result_slides_folder_id = result.files.first.id
    else
      folder_metadata = { name: "ResultSlides", mime_type: "application/vnd.google-apps.folder", parents: [ folder_id ] }
      folder = drive_service.create_file(folder_metadata, supports_all_drives: true)
      result_slides_folder_id = folder.id
    end

    # Copy the template into ResultSlides
    copied_file_metadata = { name: "PMO Report #{Time.now.to_i}", parents: [ result_slides_folder_id ] }

    copied = drive_service.copy_file(template_file_id, copied_file_metadata, supports_all_drives: true)

    presentation_id = copied.id
    Rails.logger.info("Created presentation id: #{presentation_id}")

    # 5. (Optional) Get and modify slides
    presentation = slides_service.get_presentation(presentation_id)
    Rails.logger.info("Presentation title: #{presentation.title}")

    # slide_ids = presentation.slides.map(&:object_id)
    parsed = JSON.parse(presentation.to_json)
    slide_ids = parsed["slides"].map.with_index do |slide, i|
      slide["objectId"]
    end

    title_cover_slide_id = slide_ids[TITLE_COVER_SLIDE_INDEX].to_s
    overall_summary1_slide_id = slide_ids[OVERALL_SUMMARY1_SLIDE_INDEX].to_s
    overall_summary_slide_id = slide_ids[OVERALL_SUMMARY2_SLIDE_INDEX].to_s
    overall_summary3_slide_id = slide_ids[OVERALL_SUMMARY3_SLIDE_INDEX].to_s
    cost_template_id = slide_ids[COST_SUMMARY_SLIDE_INDEX].to_s
	  cost_template2_id = slide_ids[COST_SUMMARY_SLIDE2_INDEX].to_s
    evm_template_id  = slide_ids[EVM_SUMMARY_SLIDE_INDEX].to_s
    evm_monthly_template_id  = slide_ids[EVM_MONTHLY_SLIDE_INDEX].to_s
   
    # TODO change date to today
    today = Date.today
    # today = Date.new(2025, 6, 30)
    last_month = today << 1
    this_year = today.year
 
    record = records_all.first || {}
    record_for_evm = records_month_unit.first || record

    requests = []

    date_value = records_week[0]["記入日(基準日)"]
    formatted_date = date_value.to_date.strftime("%Y年%m月%d日")

    shared_placeholders = {
    # "{{PROJECT_NAME}}" => record["Project名"],
    # "{{YEAR}}" => record["Year"],
    # "{{MONTH}}" => record["Month"],
    # "{{PM}}" => record["担当PM"],
    # "{{SPI}}" => record["SPI"],
    # "{{CPI}}" => record["CPI"],
    # "{{PV_人日}}" => record["PV_人日"],
    # "{{EV_人日}}" => record["EV_人日"],
    # "{{AC_人日}}" => record["AC_人日"],
    "{{PROJECT_COUNT}}" => records_month_unit.count,
    "{{SPI_ISSUES}}" => records_month_unit.count { |r| r["SPI"].to_f < 0.9 rescue false },
    "{{CPI_ISSUES}}" => records_month_unit.count { |r| r["CPI"].to_f < 1.0 rescue false },
    # "{{EVM_DATE}}" => today.strftime("%Y年%-m月%-d日"),
    "{{基準月}}" => "#{record["Year"]}年#{record["Month"]}月",
    "{{EVM基準日}}" => formatted_date,
    "{{作成日}}" => today.strftime("%Y年%-m月%-d日")
    }.transform_values { |v| v.nil? ? "-" : v.to_s }

    [title_cover_slide_id, overall_summary1_slide_id].each do |page_id|
      shared_placeholders.each do |ph, text|
        requests << {
          replace_all_text: {
            contains_text: { text: ph, match_case: false },
            replace_text: text.to_s,
            page_object_ids: [page_id]
          }
        }
      end
    end

    summary_placeholders = {
      "{{該当年月}}" => "#{last_month.year}年#{last_month.month}月",
      "{{取得年月日}}" => today.strftime("%Y年%-m月%-d日"),
      "{{今年度}}" => "#{this_year}年",
      "{{○回目}}" => "○回目",
      "{{取得月}}" => "#{today.month}月",
    }

    summary_placeholders.each do |ph, text|
      requests << {
        replace_all_text: {
          contains_text: { text: ph, match_case: false },
          replace_text: text.to_s,
          page_object_ids: [overall_summary_slide_id]
        }
      }
    end

    slides_service.batch_update_presentation(
      presentation_id,
      SlidesV1::BatchUpdatePresentationRequest.new(requests: requests)
    )

    presentation = slides_service.get_presentation(presentation_id)
    slide = presentation.slides[COST_SUMMARY_SLIDE_INDEX]
    parsed = JSON.parse(presentation.to_json)
    slide_ids = parsed["slides"].map.with_index do |slide, i|
      slide["objectId"]
    end
    template_slide6_id = slide_ids[COST_SUMMARY_SLIDE_INDEX]
    template_slide7_id = slide_ids[COST_SUMMARY_SLIDE2_INDEX]
    template_slide8_id = slide_ids[EVM_SUMMARY_SLIDE_INDEX]
    template_slide9_id = slide_ids[EVM_MONTHLY_SLIDE_INDEX]

    # Call Slide 5 method
    insert_and_fill_slide5_table(slides_service, presentation_id, records_week, overall_summary3_slide_id)

    presentation = slides_service.get_presentation(presentation_id)
    ppt_slides = presentation.slides.count
    total_slides = 1
    all_requests = []

    grouped_records = records_all.group_by { |r| r["取引先（企業）名"] }
    company_counter = 1

    grouped_records.each do |company, records|
      company_label = "2.#{company_counter}"

      records_for_summary = records.select { |r| r["月単位"].nil? }
      template_table6_id = find_table_id(slides_service, presentation_id, template_slide6_id)
      record_index = 0 
      records_for_summary.each_slice(8).with_index do |record_chunk, chunk_index|

      # Duplicate Slide 6 with a new table ID
      new_slide6_id = SecureRandom.uuid
      new_table6_id = SecureRandom.uuid

      all_requests << {
        duplicate_object: {
          object_id_prop: template_slide6_id,
          object_ids: {
            template_slide6_id => new_slide6_id,
            template_table6_id => new_table6_id
          }
        }
      }
      all_requests << update_slides_position_request(new_slide6_id, total_slides + ppt_slides)
      all_requests << replace_text_request("{{プロファイル名タイトル}}",  "#{company_label}. #{company}", new_slide6_id)
      all_requests <<  replace_text_request("{{担当PM}}", record["担当PM"], new_slide6_id)
       # Build table fill requests without touching the API
      all_requests.concat(insert_and_fill_table_requests(record_chunk, new_table6_id,start_index: record_index + 1))
      record_index += record_chunk.size # increment counter
      total_slides += 1   # simulate increment
    end

      # Slide 7 (loop)
      records_for_summary.each do |record|
        new_evm_project_id = SecureRandom.uuid
        all_requests.concat([
          duplicate_object_request(template_slide7_id, new_evm_project_id),
          update_slides_position_request(new_evm_project_id, total_slides + ppt_slides),
          replace_text_request("{{プロファイル名タイトル}}", "#{company_label}. #{company}", new_evm_project_id),
          replace_text_request("{{プロジェクト名}}", record["Project名"], new_evm_project_id),
          replace_text_request("{{担当PM}}", record["担当PM"], new_evm_project_id),
          replace_text_request("{{スケジュール}}", record["SPI"].to_f < 0.9 ? "遅延" : "OK", new_evm_project_id),
          replace_text_request("{{コスト}}", record["CPI"].to_f < 1 ? "経過" : "OK", new_evm_project_id)
        ])
        total_slides += 1   # increment index
      end

      # Slide 8 & 9
      projects = records.group_by { |r| r["Project名"] }
      projects.each do |project_name, project_records|
        record_no_month   = project_records.find { |r| r["月単位"].nil? }
        record_with_month = project_records.find { |r| !r["月単位"].nil? }

        if record_no_month
          new_evm_summary_id = SecureRandom.uuid
          requests_slide8 = [
            duplicate_object_request(template_slide8_id, new_evm_summary_id),
            update_slides_position_request(new_evm_summary_id, total_slides + ppt_slides),
            replace_text_request("{{プロファイル名タイトル}}", "#{company_label}. #{company}", new_evm_summary_id),
            replace_text_request("{{プロジェクト名}}", record_no_month["Project名"], new_evm_summary_id),
            replace_text_request("{{担当PM}}", record_no_month["担当PM"], new_evm_summary_id)
          ]

          all_requests.concat(requests_slide8)
          all_requests.concat(fill_evm_slides(slides_service, presentation_id, record_no_month, new_evm_summary_id))
          total_slides += 1    # increment index
          Rails.logger.info("Already complete duplicate Slide 8--->  #{project_name}")
        end
        if record_with_month
          new_evm_monthly_id = SecureRandom.uuid

          requests_slide9 = [
            duplicate_object_request(template_slide9_id, new_evm_monthly_id),
            update_slides_position_request(new_evm_monthly_id, total_slides + ppt_slides),
            replace_text_request("{{プロファイル名タイトル}}", "#{company_label}. #{company}", new_evm_monthly_id),
            replace_text_request("{{プロジェクト名}}", record_with_month["Project名"], new_evm_monthly_id),
            replace_text_request("{{担当PM}}", record_with_month["担当PM"], new_evm_monthly_id)
          ]

          all_requests.concat(requests_slide9)
          all_requests.concat(
            fill_evm_slides(slides_service, presentation_id, record_with_month, new_evm_monthly_id)
          )
          total_slides += 1  # increment index
          Rails.logger.info("Already complete duplicate Slide 9--->  #{project_name}")
        end
      end
       # increment once per company
     company_counter += 1
    end

    # 🔥 Final single API call for slide 6,7,8,9
    slides_service.batch_update_presentation(
        presentation_id,
        Google::Apis::SlidesV1::BatchUpdatePresentationRequest.new(requests: all_requests)
    )
    
    Rails.logger.info("Already complete--->")
    
    # Optional: Delete original template slide to keep it clean
    delete_request = [
      {
        delete_object: {
          object_id_prop: template_slide6_id
        }
      },
      {
        delete_object: {
          object_id_prop: template_slide7_id
        }
      },
      {
        delete_object: {
          object_id_prop: template_slide8_id
        }
      },
      {
        delete_object: {
          object_id_prop: template_slide9_id
        }
      }
    ]
    slides_service.batch_update_presentation(
      presentation_id,
      Google::Apis::SlidesV1::BatchUpdatePresentationRequest.new(requests: delete_request)
    )

    Rails.logger.warn("end code --")
    export_url = "https://docs.google.com/presentation/d/#{presentation_id}/export/pptx"
    pptx_path = Rails.root.join("tmp", "generated_slide_#{Time.now.to_i}.pptx")
    access_token = drive_service.authorization.fetch_access_token!["access_token"]

    URI.open(export_url, "Authorization" => "Bearer #{access_token}") do |remote_file|
      File.open(pptx_path, "wb") { |file| file.write(remote_file.read) }
    end

    pptx_path.to_s
  end

  def self.duplicate_object_request(template_id, new_id)
    {
      duplicate_object: {
        object_id_prop: template_id,
        object_ids: { template_id => new_id }
      }
    }
  end

  def self.update_slides_position_request(slide_id, index)
    {
      update_slides_position: {
        slide_object_ids: [slide_id],
        insertion_index: index
      }
    }
  end

  def self.replace_text_request(placeholder, value, slide_id)
    {
      replace_all_text: {
        contains_text: { text: placeholder, match_case: true },
        replace_text: value.to_s,
        page_object_ids: [slide_id]
      }
    }
  end

  def self.insert_and_fill_table_requests(records, new_table_id, start_index: 1)
    number_of_rows = records.size

    insert_requests = [
      {
        insert_table_rows: {
          table_object_id: new_table_id,
          cell_location: { row_index: 0 },
          insert_below: true,
          number: number_of_rows
        }
      }
    ]

    background_requests = [
      {
        update_table_cell_properties: {
          object_id_prop: new_table_id,
          table_range: {
            location:   { row_index: 1, column_index: 0 },
            row_span:   number_of_rows,
            column_span: 9
          },
          table_cell_properties: {
            table_cell_background_fill: {
              solid_fill: { color: { rgb_color: { red: 1.0, green: 1.0, blue: 1.0 } } }
            }
          },
          fields: "tableCellBackgroundFill.solidFill.color"
        }
      }
    ]

    fill_requests = records.each_with_index.flat_map do |record, idx|
      r = idx + 1
      display_num = start_index + idx # 👈 continuous numbering
      [
        { insert_text: { object_id_prop: new_table_id, cell_location: { row_index: r, column_index: 0 }, text: display_num.to_s } },
        { insert_text: { object_id_prop: new_table_id, cell_location: { row_index: r, column_index: 1 }, text: record["Project名"].to_s } },
        { insert_text: { object_id_prop: new_table_id, cell_location: { row_index: r, column_index: 7 }, text: record["SPI"].to_f < 0.9 ? "遅延" : "OK" } },
        { insert_text: { object_id_prop: new_table_id, cell_location: { row_index: r, column_index: 8 }, text: record["CPI"].to_f < 1 ? "経過" : "OK" } }
      ]
    end

    insert_requests + background_requests + fill_requests
  end

  def self.find_table_id(slides_service, presentation_id, template_slide6_id)
    presentation = slides_service.get_presentation(presentation_id)
    slide = presentation.slides.find { |s| s.object_id_prop == template_slide6_id }
    raise "Template slide not found: #{template_slide6_id}" unless slide

    table_element = slide.page_elements.find(&:table)
    raise "No table found in Slide #{template_slide6_id}" unless table_element

    table_element.object_id_prop
  end

  def self.insert_and_fill_slide5_table(slides_service, presentation_id, records_week, slide5_id)
    # Validate the base slide exists
    pres = slides_service.get_presentation(presentation_id)
    base_slide = pres.slides.find { |s| s.object_id_prop == slide5_id }
    raise "Slide 5 not found: #{slide5_id}" unless base_slide

    # Build per-company data rows
    grouped = records_week.group_by { |r| r["取引先（企業）名"] }
    data_rows = []
    seq = 1
    grouped.each do |company, recs|
      prj_count = recs.map { |r| r["Project名"] }.uniq.count
      spi_violation_count = recs.count { |r| r["SPI"].to_f < 0.9 rescue false }
      cpi_violation_count = recs.count { |r| r["CPI"].to_f < 1.0 rescue false }

      data_rows << {
        no: seq,
        company: company,
        exec: recs.first["担当執行役員"],
        super_pm: recs.map { |r| r["担当統括PM"] }.compact.uniq.join("\n"),
        pm: recs.map{|r| r["担当PM"] }.compact.uniq.join("\n"),
        prj_count: prj_count,
        spi_violation: spi_violation_count > 0 ? "#{spi_violation_count}件違反" : "0",
        cpi_violation: cpi_violation_count > 0 ? "#{cpi_violation_count}件違反" : "0"
      }
      seq += 1
    end

    # Paginate: 10 data rows per slide
    pages = data_rows.each_slice(8).to_a
    final_page_index = pages.size - 1

    # Duplicate Slide 5 template for additional pages
    pres = slides_service.get_presentation(presentation_id)
    parsed = JSON.parse(pres.to_json)
    slide_id_list = parsed["slides"].map { |s| s["objectId"] }
    base_index = slide_id_list.index(slide5_id) || 0

    page_slide_ids = [ slide5_id ]
    if pages.size > 1
      dup_requests = []
      (1...pages.size).each do |i|
        new_id = SecureRandom.uuid
        page_slide_ids << new_id
        dup_requests << {
          duplicate_object: {
            object_id_prop: slide5_id,
            object_ids: { slide5_id => new_id }
          }
        }
        dup_requests << {
          update_slides_position: {
            slide_object_ids: [ new_id ],
            insertion_index: base_index + i
          }
        }
      end
      slides_service.batch_update_presentation(
        presentation_id,
        Google::Apis::SlidesV1::BatchUpdatePresentationRequest.new(requests: dup_requests)
      )
    end

    # Calculate totals once
    total_prj_count = data_rows.sum { |r| r[:prj_count].to_i }
    total_spi_violations = data_rows.sum { |r| r[:spi_violation].to_s.gsub("件違反", "").to_i }
    total_cpi_violations = data_rows.sum { |r| r[:cpi_violation].to_s.gsub("件違反", "").to_i }
    percent_spi = total_prj_count > 0 ? (total_spi_violations.to_f / total_prj_count * 100).round : 0
    percent_cpi = total_prj_count > 0 ? (total_cpi_violations.to_f / total_prj_count * 100).round : 0

    # Fill each slide
    page_slide_ids.each_with_index do |target_slide_id, page_idx|
      pres_current = slides_service.get_presentation(presentation_id)
      target_slide = pres_current.slides.find { |s| s.object_id_prop == target_slide_id }
      raise "Slide not found during pagination fill: #{target_slide_id}" unless target_slide

      table_elements = target_slide.page_elements.select(&:table)
      raise "No tables found in Slide 5 page #{page_idx + 1}" if table_elements.empty?

      table = table_elements.first
      table_id = table.object_id_prop

      # Build rows for this page
      rows_for_page = pages[page_idx].dup

      # Append calculation rows only on the final slide
      if page_idx == final_page_index
        rows_for_page << {
          no: "", company: "", exec: "", super_pm: "", pm: "",
          prj_count: total_prj_count,
          spi_violation: total_spi_violations,
          cpi_violation: total_cpi_violations
        }
        rows_for_page << {
          no: "", company: "", exec: "", super_pm: "", pm: "",
          prj_count: "-",
          spi_violation: "#{percent_spi}%",
          cpi_violation: "#{percent_cpi}%"
        }
      end

      num_rows = rows_for_page.size
      
      insert_requests = [
        {
          insert_table_rows: {
            table_object_id: table_id,
            cell_location: { row_index: 0 },
            insert_below: true,
            number: num_rows
          }
        }
      ]

      background_requests = [
        {
          update_table_cell_properties: {
            object_id_prop: table_id,
            table_range: {
              location: { row_index: 1, column_index: 0 },
              row_span: num_rows,
              column_span: 8
            },
            table_cell_properties: {
              table_cell_background_fill: {
                solid_fill: { color: { rgb_color: { red: 1.0, green: 1.0, blue: 1.0 } } }
              }
            },
            fields: "tableCellBackgroundFill.solidFill.color"
          }
        }
      ]

      fill_requests = rows_for_page.each_with_index.map do |row, idx|
        ri = idx + 1
        [
          { insert_text: { object_id_prop: table_id, cell_location: { row_index: ri, column_index: 0 }, text: row[:no].to_s } },
          { insert_text: { object_id_prop: table_id, cell_location: { row_index: ri, column_index: 1 }, text: row[:company].to_s } },
          { insert_text: { object_id_prop: table_id, cell_location: { row_index: ri, column_index: 2 }, text: row[:exec].to_s } },
          { insert_text: { object_id_prop: table_id, cell_location: { row_index: ri, column_index: 3 }, text: row[:super_pm].to_s } },
          { insert_text: { object_id_prop: table_id, cell_location: { row_index: ri, column_index: 4 }, text: row[:pm].to_s } },
          { insert_text: { object_id_prop: table_id, cell_location: { row_index: ri, column_index: 5 }, text: row[:prj_count].to_s } },
          { insert_text: { object_id_prop: table_id, cell_location: { row_index: ri, column_index: 6 }, text: row[:spi_violation].to_s } },
          { insert_text: { object_id_prop: table_id, cell_location: { row_index: ri, column_index: 7 }, text: row[:cpi_violation].to_s } }
        ]
      end.flatten

      slides_service.batch_update_presentation(
        presentation_id,
        Google::Apis::SlidesV1::BatchUpdatePresentationRequest.new(
          requests: insert_requests + background_requests + fill_requests
        )
      )
    end
  end


  def self.fill_evm_slides(slides_service, presentation_id, record, evm_template_id)
    if record["SV結果の解釈\n【スケジュール差異】"].to_s.include?("大幅なスケジュール遅延が発生しています")
      sv = "大幅なスケジュール遅延が発生しています。\n計画より 約#{record["SV"].to_f.abs}人日分の作業が遅れている。"
    elsif record["SV結果の解釈\n【スケジュール差異】"].to_s.include?("計画通り進行しています")
      sv = "計画通り進行しています。\n#{record["SV"].to_f.abs}人日分の遅れ。"
    elsif record["SV結果の解釈\n【スケジュール差異】"].to_s.include?("やや遅れています（軽微な遅延）")
      sv = "やや遅れています（軽微な遅延）。\n計画より 約#{record["SV"].to_f.abs}人日分の作業が遅れている。"
    elsif record["SV結果の解釈\n【スケジュール差異】"].to_s.include?("やや前倒しで順調に進行しています")
      sv = "やや前倒しで順調に進行しています。"
    elsif record["SV結果の解釈\n【スケジュール差異】"].to_s.include?("大幅に前倒しで進行しています")
      sv = "大幅に前倒しで進行しています"
    end
  
    if record["CV結果の解釈\n【コスト差異】"].to_s.include?("コスト超過が大きく、改善が必要です")
      cv = "コスト超過が大きく、改善が必要です。\n想定より 約#{record["CV"].to_f.abs}人日分の工数を超過している。"
    elsif record["CV結果の解釈\n【コスト差異】"].to_s.include?("やや効率的に進んでいます")
      cv = "やや効率的に進んでいます。\n約#{record["CV"].to_f.abs}人日分のコスト節約（予算より少ない工数で成果が出ている）。"
    elsif record["CV結果の解釈\n【コスト差異】"].to_s.include?("ややコスト超過しています")
      cv = "ややコスト超過しています。\n想定より 約#{record["CV"].to_f.abs}人日分の工数を超過している。"
    elsif record["CV結果の解釈\n【コスト差異】"].to_s.include?("コスト効率が非常に良好です")
      cv = "コスト効率が非常に良好です。\n約#{record["CV"].to_f.abs}人日分の工数が節約されている（予定より少ない工数で済んでいる）。"
    elsif record["CV結果の解釈\n【コスト差異】"].to_s.include?("ほぼ計画通りのコスト進行です")
      if record["CV"].to_s.include?("-")
        cv = "ほぼ計画通りのコスト進行です。\n想定より 約#{record["CV"].to_f.abs}人日分の工数を超過している。"
      else
        cv = "ほぼ計画通りのコスト進行です。\n#{record["CV"].to_f.abs}人日分の工数が節約されている。"
      end  
    end

    if record["SPI結果の解釈\n【進捗効率の健全性評価】"].to_s.include?("大幅なスケジュール遅延が発生しています")
      spi_value = record["SPI"].to_f.abs
      spi_percentage = (spi_value * 100).round(1)
      spi = "大幅なスケジュール遅延が発生しています。\n進捗は 計画の約#{spi_percentage}% のスピードで進んでいる。"
    elsif record["SPI結果の解釈\n【進捗効率の健全性評価】"].to_s.include?("ややスケジュール遅れがあります")
      spi_value = record["SPI"].to_f.abs
      spi_percentage = (spi_value * 100).round(1)
      spi = "ややスケジュール遅れがあります\n予定の#{spi_percentage}%のスピードで進行。"
    elsif record["SPI結果の解釈\n【進捗効率の健全性評価】"].to_s.include?("予定通りに進んでいます")
      spi = "予定通りに進んでいます。\nスケジュール通りのスピードで作業が進んでいる。"
    elsif record["SPI結果の解釈\n【進捗効率の健全性評価】"].to_s.include?("スケジュール遅れが目立ちます")
      spi_value = record["SPI"].to_f.abs
      spi_percentage = (spi_value * 100).round(1)
      spi = "スケジュール遅れが目立ちます。\n作業進捗は計画の#{spi_percentage}%。"
    elsif record["SPI結果の解釈\n【進捗効率の健全性評価】"].to_s.include?("スケジュールより前倒しで順調に進んでいます")
      spi = "スケジュールより前倒しで順調に進んでいます。"
    end

    if record["CPI結果の解釈\n【コスト効率の健全性評価】"].to_s.include?("コスト超過が大きく、対策が必要です")
      cpi_value = record["CPI"].to_f.abs
      cpi_percentage = (cpi_value * 100).round(1)
      cpi = "コスト超過が大きく、対策が必要です。\n投入した工数のうち、#{cpi_percentage}%分の成果しか得られていない。 "
    elsif record["CPI結果の解釈\n【コスト効率の健全性評価】"].to_s.include?("非常に効率的に作業が進められています")
      cpi_value = record["CPI"].to_f.abs
      cpi_percentage = (cpi_value * 100).round(1)
      cpi = "非常に効率的に作業が進められています。\n投入した工数の#{cpi_percentage}%分の成果が出ている。"
    elsif record["CPI結果の解釈\n【コスト効率の健全性評価】"].to_s.include?("コスト効率良好です")
      cpi_value = record["CPI"].to_f.abs
      cpi_percentage = (cpi_value * 100).round(1)
      cpi = "コスト効率良好です。\n投入コストに対して#{cpi_percentage}%の成果が得られている。"
    elsif record["CPI結果の解釈\n【コスト効率の健全性評価】"].to_s.include?("ほぼ計画通りにコスト管理されています")
      cpi_value = record["CPI"].to_f.abs
      cpi_percentage = (cpi_value * 100).round(1)
      cpi = "ほぼ計画通りにコスト管理されています。\n投入したコストのうち、#{cpi_percentage}%が成果に結びついている。"
    elsif record["CPI結果の解釈\n【コスト効率の健全性評価】"].to_s.include?("ややコスト超過傾向があります")
      cpi_value = record["CPI"].to_f.abs
      cpi_percentage = (cpi_value * 100).round(1)
      cpi = "ややコスト超過傾向があります。\n投入した工数のうち、#{cpi_percentage}%分の成果のみ出ている。"
    end

    placeholders = {
      "{{SV}}" => record["SV"].to_s,
      "{{CV}}" => record["CV"].to_s,
      "{{SPI}}" => record["SPI"].to_s,
      "{{CPI}}" => record["CPI"].to_s,
      "{{PV_人日}}" => record["PV_人日"].to_s,
      "{{EV_人日}}" => record["EV_人日"].to_s,
      "{{AC_人日}}" => record["AC_人日"].to_s,
      "{{SV評価}}" => sv,
      "{{CV評価}}" => cv,
      "{{SPI評価}}" => spi,
      "{{CPI評価}}" => cpi
    }

    placeholders.map do |placeholder, value|
      {
        replace_all_text: {
          contains_text: { text: placeholder },
          replace_text: value,
          page_object_ids: [evm_template_id]
        }
      }
    end

    # slides_service.batch_update_presentation(
    #   presentation_id,
    #   Google::Apis::SlidesV1::BatchUpdatePresentationRequest.new(requests: requests)
    # )
  end

  def self.group_records_by_condition(records_all)
    today = Date.today
    # today = Date.new(2025, 6, 30)
    last_month = today << 1
    last_month_str = "#{last_month.month}月"

    week = []
    month_unit = []

    records_all.each do |r|
      next unless r.is_a?(Hash) 
      if r["月単位"].to_s.strip == last_month_str
        month_unit << r
      else
        week << r
      end
    end

    { all: records_all, week: week, month_unit: month_unit }
  end

end