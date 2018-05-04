require 'roo'
require 'countries'
require 'going_postal'

# Pulls out specific columns from the E-mail sheet for MailChimp
class Newsletter
  attr_reader :sheet_names,
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
    @sheet_names = ['Email', 'E-Mail', 'E-mail', 'Electronic']
    @send_past_due = true
  end

  # Generate csv file with the specified columns
  def generate
    CSV.open(output_filename, 'w') do |csv|
      sheet_names.each do |sheet_name|
        sheet(sheet_name).each_with_index do |row, index|
          if index.zero?
            # headers, restore original column name
            csv << row.keys.map { |value| XLS_COLUMNS[value] }
          end
          unless row[:email].respond_to?(:split)
            puts "Email skipped because it is empty"
            next
          end

          original_email = row[:email].dup
          original_email.split.each do |email|
            next if past_due?(row[:member_through])
            puts "Email reformatted from: [#{original_email}] to [#{email}]" if email != row[:email]
            unless valid_email?(email)
              puts "Email invalid: #{email}"
              next
            end
            row[:email] = email

            unless row[:member_through].respond_to?(:strftime)
              new_date = Date.parse("01-#{row[:member_through]}")
              puts "Correcting date: #{row[:member_through]} to #{new_date.strftime('%b-%y')}"
              row[:member_through] = new_date
            end
            row[:member_through] = row[:member_through].strftime('%b-%y')
            row[:zip] = format_zipcode(row[:zip], row[:country])

            csv << row.values.map { |value| clean_html(value) }
          end
        end
      end
    end
  end

  private

  def past_due?(date)
    return false if @send_past_due == true
    today = Date.today
    date < Date.new(today.year, today.month - 1, 1)
  end

  # true if it contains an @ symbol and . after the @ symbol
  def valid_email?(email)
    return false if email.nil? || email.empty?
    email.include?('@') && email.split('@')[1].include?('.')
  end

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

  def sheet(sheet_name)
    ::Roo::Spreadsheet
      .open(file_path, clean: true)
      .sheet(sheet_name)
      .parse(XLS_COLUMNS)
  rescue RangeError => e
    puts "Failed to open #{sheet_name}: #{e.message}"
    []
  end
end
