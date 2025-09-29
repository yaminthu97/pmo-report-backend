class PowerPointGeneratorJob < ApplicationJob
  queue_as :default

  def perform_old(xlsx_path, job_id, google_token)
    spreadsheet = Roo::Spreadsheet.open(xlsx_path)
    sheet_name = spreadsheet.sheets.first

    spreadsheet.default_sheet = sheet_name
    raw_headers = spreadsheet.row(4).map(&:to_s)

    headers = map_duplicate_headers(raw_headers)
    all_records = extract_filtered_records(spreadsheet, headers)
    grouped_records = ExportPowerPointSchedulerModule.group_records_by_condition(all_records)
    today = Date.today
    # today = Date.new(2025, 6, 30)
    last_month_date = today << 1
    target_year = last_month_date.year
    target_month = last_month_date.month
    target_week = Date.new(target_year, target_month, -1).cweek

    filtered_all = grouped_records[:all].select do |r|
      (normalize_number(r["Year"]) == target_year &&
      normalize_number(r["Month"]) == target_month &&
      normalize_number(r["Week"]) == target_week) ||
        r["月単位"].to_s.strip == "#{target_month}月"
    end

    records_week = filtered_all.select do |r|
      normalize_number(r["Week"]) == target_week
    end

    records_month = filtered_all.select do |r|
      r["月単位"].to_s.strip == "#{target_month}月"
    end

    combined_records = records_week + records_month
    if combined_records.empty?
      Rails.logger.warn("No matching records found. Skipping export.")
      return
    end

    grouped_final = {
      all: combined_records,
      week: records_week,
      month_unit: records_month
    }

    pptx_path = ExportPowerPointSchedulerModule.export_excel_data(grouped_final, google_token)
    final_path = Rails.root.join("tmp", "generated_slide_#{job_id}.pptx")
    FileUtils.mv(pptx_path, final_path)
    File.delete(xlsx_path) if File.exist?(xlsx_path)
  end

  def perform(xlsx_path, job_id, google_token)
    spreadsheet = Roo::Spreadsheet.open(xlsx_path)
    sheet_name = spreadsheet.sheets.first

    spreadsheet.default_sheet = sheet_name
    raw_headers = spreadsheet.row(4).map(&:to_s)

    headers = map_duplicate_headers(raw_headers)
    all_records = extract_filtered_records(spreadsheet, headers)
    grouped_records = ExportPowerPointSchedulerModule.group_records_by_condition(all_records)
    today = Date.today
    # today = Date.new(2025, 6, 30)
    last_month_date = today << 1
    target_year = last_month_date.year
    target_month = last_month_date.month
    target_week = Date.new(target_year, target_month, -1).cweek

    filtered_all = grouped_records[:all].select do |r|
      (normalize_number(r["Year"]) == target_year &&
      normalize_number(r["Month"]) == target_month &&
      normalize_number(r["Week"]) == target_week) ||
        r["月単位"].to_s.strip == "#{target_month}月"
    end

    records_week = filtered_all.select do |r|
      normalize_number(r["Week"]) == target_week
    end

    records_month = filtered_all.select do |r|
      r["月単位"].to_s.strip == "#{target_month}月"
    end

    combined_records = records_week + records_month
    if combined_records.empty?
      Rails.logger.warn("No matching records found. Skipping export.")
      return
    end

    grouped_final = {
      all: combined_records,
      week: records_week,
      month_unit: records_month
    }

    pptx_path = ExportPowerPointSchedulerModule.export_excel_data(grouped_final, google_token)
    final_path = Rails.root.join("tmp", "generated_slide_#{job_id}.pptx")
    FileUtils.mv(pptx_path, final_path)
    File.delete(xlsx_path) if File.exist?(xlsx_path)
  end

  private

  def extract_filtered_records(spreadsheet, headers)
    (5..spreadsheet.last_row).map do |i|
      values = spreadsheet.row(i)

      row = Hash[headers.zip(values)].reject { |k, _| k.nil? || k.strip.empty? }

      excel_date = row["記入日(基準日)"]
      if excel_date.is_a?(Numeric)
        base_date = Date.new(1899, 12, 30)
        row["記入日(基準日)"] = (base_date + excel_date.to_i).strftime("%Y/%m/%d")
      end

      next if row["Year"].to_s.strip.empty?

      begin
        row["SV"] = row["SV"].to_f.round(1) if row["SV"]
        row["CV"] = row["CV"].to_f.round(1) if row["CV"]
        row["SPI"] = row["SPI"].to_f.round(1) if row["SPI"]
        row["CPI"] = row["CPI"].to_f.round(2) if row["CPI"]

        pv_m = row["PV_人月"].to_f
        ev_m = row["EV_人月"].to_f
        ac_m = row["AC_人月"].to_f

        pv_d = row["PV_人日"].to_f
        ev_d = row["EV_人日"].to_f
        ac_d = row["AC_人日"].to_f

        pv_h = row["PV_人時間"].to_f
        ev_h = row["EV_人時間"].to_f
        ac_h = row["AC_人時間"].to_f

        row["PV_人月"], row["EV_人月"], row["AC_人月"] = pv_m.round(1), ev_m.round(1), ac_m.round(1)
        row["PV_人日"], row["EV_人日"], row["AC_人日"] = pv_d.round(1), ev_d.round(1), ac_d.round(1)
        row["PV_人時間"], row["EV_人時間"], row["AC_人時間"] = pv_h.round(1), ev_h.round(1), ac_h.round(1)
      rescue => e
        Rails.logger.warn("Row ##{i} caused an error: #{e.message}")
        next
      end

      row
    end.compact
  end
  def map_duplicate_headers(headers)
    header_map = {
      "PV"  => "PV_人月",
      "EV"  => "EV_人月",
      "AC"  => "AC_人月",
      "PV2" => "PV_人日",
      "EV3" => "EV_人日",
      "AC4" => "AC_人日",
      "PV5" => "PV_人時間",
      "EV6" => "EV_人時間",
      "AC7" => "AC_人時間"
    }

    headers.map { |h| header_map[h] || h }
  end
 def normalize_number(val)
    val.to_s.gsub(/[^0-9]/, "").to_i
  end
end
  