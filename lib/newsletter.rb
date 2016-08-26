require 'roo'
require 'countries'
require 'going_postal'

# Pulls out specific columns from the E-mail sheet for MailChimp
class Newsletter
  attr_reader :sheet_name,
              :file_path,
              :output_path,
              :filename,
              :output_format

  DEFAULT_COUNTRY = 'US'.freeze

  XLS_COLUMNS = {
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
  }.freeze

  def initialize(file_path, filename = nil)
    @file_path = file_path
    @filename = filename || File.basename(file_path, File.extname(file_path))
    @output_format = 'csv'
    @output_path = 'output'
    @sheet_name = 'E-mail'
  end

  # Generate csv file with the specified columns
  def generate
    CSV.open(output_filename, 'w') do |csv|
      sheet.each_with_index do |row, index|
        unless index.zero?
          row[:member_through] = row[:member_through].strftime('%b-%y')
          row[:zip] = format_zipcode(row[:zip], row[:country])
        end
        csv << row.values.map { |value| clean_html(value) }
      end
    end
  end

  private

  def country_alpha2(country)
    return if country.nil?
    cntry = ISO3166::Country.new(country) ||
            ISO3166::Country.find_country_by_alpha3(country) ||
            ISO3166::Country.find_country_by_name(country)
    if cntry.nil?
      puts "Invalid country: #{country}"
      return country
    end
    cntry.alpha2
  end

  def format_zipcode(value, country)
    return if value.nil?
    country_alpha2 = country_alpha2(country)
    puts "Defaulting Country to #{DEFAULT_COUNTRY} for postal code: #{value}" if country_alpha2.nil?

    postal = postal_code_for(value, country_alpha2)
    return postal if postal

    puts "Invalid postal code: #{value} for #{country} [#{country_alpha2}]"
    value
  end

  def postal_code_for(value, country_alpha2)
    GoingPostal.postcode?(value, country_alpha2 || DEFAULT_COUNTRY)
  end

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
      .parse(XLS_COLUMNS)
  end
end
