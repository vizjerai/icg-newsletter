class Newsletter
  attr_reader :sheet_name, :file_path,
    :xls_columns,
    :output_path, :filename, :output_format

  def initialize(file_path, filename = nil)
    @file_path = file_path
    @filename = filename || File.basename(file_path, File.extname(file_path))
    @output_format = 'csv'
    @output_path = 'output'
    @sheet_name = 'E-mail'
    @xls_columns = {
      chapter: 'Chapter (Primary)',
      member_through: 'Member Thru',
      last_name: 'Last Name',
      first_name: 'First Name',
      address: 'Address',
      city: 'City',
      state: 'State',
      zip: 'Zip',
      country: 'Cntry',
      email: 'Email'
    }
  end

  # Generate csv file with the specified columns
  def generate
    require 'roo'

    CSV.open(output_filename, 'w') do |csv|
      sheet.each_with_index do |row, index|
        unless index.zero?
          row[:member_through] = row[:member_through].strftime('%b-%y')
        end
        csv << row.values.map do |value|
          clean_html(value)
        end
      end
    end
  end

  private

  def clean_html(value)
    return if value.nil?
    value.to_s.gsub(/<\/?[^>]*>/, '')
  end

  def output_filename
    File.join(output_path, "#{filename}.#{output_format}")
  end

  def sheet
    ::Roo::Spreadsheet
      .open(file_path, clean: true)
      .sheet(sheet_name)
      .parse(xls_columns)
  end
end
