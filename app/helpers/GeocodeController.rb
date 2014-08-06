require 'pathname'
require 'spreadsheet'
require 'geocoder'

#*******GEOCODER INFO************
# Site: https://github.com/alexreisner/geocoder#google-google-google_premier
# Used only free services that cover world.
#
# :google         2500/day              7/10/14
# :bing           50000/day             7/10/14
# :nominatim      1/sec                 7/10/14
# :yandex         25000/day             7/10/14
#*******GEOCODER INFO************

class GeocodeController
  BUFFER_SIZE = 1000
  FIXNUM_MAX = (2**(0.size * 8 -2) -1)
  @@keys = {}
  
  def initialize(row, col, book_in, filename)
  	@filename = filename
    @utils = []
    @key_order = []
    @geocoded_vals = []
    @row = row - 1
    @col = col - 1
    @book_in = book_in
    @sheet_in = @book_in.worksheet(0)
    @curr_key = nil
    
    if(@@keys.length == 0)
      read_keys
    end
  end
    
  #*******CREATE_UTILS************
  # Create an array of GeocodeUtils for multi-threading
  # for faster geocoding
  #*******CREATE_UTILS************

  def create_utils
    vals = []
    vals = @sheet_in.column(@col).to_a
    vals = vals[@row...vals.length]
    GeocodeUtil.set_total_vals(vals.length)
    
    if(vals.length % BUFFER_SIZE == 0)
      num_utils = vals.length/BUFFER_SIZE
    else
      num_utils = vals.length/BUFFER_SIZE + 1
    end
    
    (0...num_utils).each do |i|
      if(i < num_utils - 1)
        @utils.push(GeocodeUtil.new(vals[i * BUFFER_SIZE... i * BUFFER_SIZE + BUFFER_SIZE]))
      else
        @utils.push(GeocodeUtil.new(vals[i * BUFFER_SIZE...vals.length + 1]))
      end
    end
  end
  
  #*******READ_KEYS************
  # read_keys method reads the key hash from the file geo_srv_info
  # the key hash contains values for: time of next count reset, remaining queries, and API key
  # if geo_srv_info does not exist the method returns default values for the hash
  #*******READ_KEYS************

  def read_keys
    file_to_read = "#{File.dirname(__FILE__)}/geo_srv_info"
    
    if(!File.exists?(file_to_read))
      @@keys['google'] = [nil, 2500, (DateTime.now + 1).iso8601]
      @@keys['bing'] = ['AmBTPstLJhUmi10CXRQpPXneu5ToFdxidu0EWBYoxJIHFpqDN-QF9AaTkGPHkpR-', 50000, (DateTime.now + 1).iso8601]
      @@keys['nominatim'] = [nil, FIXNUM_MAX, DateTime.now.iso8601]
      @@keys['yandex'] = [nil, 25000, (DateTime.now + 1).iso8601]
      write_keys
    else
      File.open(file_to_read, 'r') do |file|
        @@keys = JSON.load(file)
        check_keys
      end
    end
    
    sort_keys
    GeocodeUtil.keys(@@keys)
  end
  
  #*******CHECK_KEYS************
  # Check timestamp on keys to determine whether
  # or not to reset query count for a key
  #*******CHECK_KEYS************

  def check_keys
  	now = DateTime.now.iso8601
    if(DateTime.strptime(@@keys['google'][2].to_s) < now)
      @@keys['google'][2] = (DateTime.now + 1).iso8601
      @@keys['google'][1] = 25000
    end
    
    if(DateTime.strptime(@@keys['bing'][2].to_s) < now)
      @@keys['bing'][2] = (DateTime.now + 1).iso8601
      @@keys['bing'][1] = 50000
    end
    
    if(DateTime.strptime(@@keys['yandex'][2].to_s) < now)
      @@keys['yandex'][2] = (DateTime.now + 1).iso8601
      @@keys['yandex'][1] = 25000
    end
  end
  
  def choose_service
    @@keys.each do |key|
      if(GeocodeUtil.total_vals < key[1][1] && key[1][1] < 50001)
        @curr_key = key[0]
        return true
      end
    end
 
    return false
  end

  #*******WRITE_KEYS************
  # write_keys method writes the key hash to the file .geo_srv_info
  # checks to see if keys with quotas can reset quotas
  #*******WRITE_KEYS************
  
  def write_keys
    file_to_write = "#{File.dirname(__FILE__)}/geo_srv_info"

    File.open(file_to_write, 'w') do |file|
      check_keys
      file.write(@@keys.to_json)
    end
  end

  def sort_keys
    temp = @@keys.sort_by { |name, info| info[1] }

    (0...temp.length).each do |i|
      @key_order.push(temp[i][0])   
    end
    
    temp = @key_order[@key_order.length - 1]
    @key_order.delete(temp)
    @key_order.unshift(temp)
  end
  
  def geocode
    create_utils
    threads = []
    
    if(choose_service)
      @utils.each do |geo_util|
        threads.push(Thread.new{
          @geocoded_vals.push(geo_util.geocode_one(@curr_key))
        })  
      end
    else
      @utils.each do |geo_util|
        threads.push(Thread.new{
          @geocoded_vals.push(geo_util.geocode_mult(@key_order))
        })  
      end       
    end
    
    threads.each do |thread|
      thread.join
    end
    
    write_keys
    write_out_sheet
  end

  def write_out_sheet
    copy_sheet
    write_vals
  end
  
  def write_vals
    curr_row = 0
    @sheet_out[curr_row, @lat_col] = "Latitude"
    @sheet_out[curr_row, @long_col] = "Longitude"
    curr_row += 1
    
    @geocoded_vals.each do |arr|
      arr.each do |val|
        @sheet_out[curr_row, @lat_col] = val[0]
        @sheet_out[curr_row, @long_col] = val[1]
        curr_row += 1
      end
    end
    
    
	@book_out.write("#{Dir.pwd}/tmp/books/#{@filename.split('.')[0]}_geocoded.xls")
  end
  
  def copy_sheet
    @book_out = Spreadsheet::Workbook.new
    @sheet_out = @book_out.create_worksheet

    row_idx = 0
    col_idx = 0
    
    @sheet_in.each do |row|
      col_idx = 0
      row.each do |col|
        @sheet_out[row_idx, col_idx] = @sheet_in[row_idx, col_idx]
        col_idx += 1
      end
      row_idx += 1
    end
    
    @lat_col = col_idx 
    @long_col = col_idx + 1
  end
  
  private :read_keys, :sort_keys
  
  class GeocodeUtil
  @@total_vals = 0
  @@sum_of_rates = 0
  @@geocoded_count = 0
  @@rate_count = 0
  @@vals_left = nil
  @@keys = nil
  @@current_key = nil
  @@avg_rate = nil
  @@const_start_time = nil
  @@last_print_time = nil
  @@first_print = true
  
  Geocoder::Configuration.always_raise << Geocoder::OverQueryLimitError
  
  def print_info
    @@geocoded_count += 1
    @@rate_count += 1
    end_time = Time.now
    elapsed = "#{Time.at(end_time - @@const_start_time).gmtime.strftime('%R:%S')}".split(":")
    rate = (@@geocoded_count/((end_time - @@const_start_time)/60))
    @@sum_of_rates += rate
    @@avg_rate = @@sum_of_rates/@@rate_count
    @@keys[@@current_key][1] -= 1
    
    if @@avg_rate > 0
      time_left = "#{Time.at((@@vals_left/@@avg_rate) * 60).gmtime.strftime('%R:%S')}".split(":")
    else
      time_left = "inf"
    end
    
    if(@@first_print || (end_time - @@last_print_time >= 1))
      @@last_print_time = Time.now
      @@first_print = false
      
      puts "\033c"
      puts "------------------------------------------------------"
      puts "Geocoding #{@@vals_left} entries."
      puts "Elapsed time: #{elapsed[0]} hours, #{elapsed[1]} minutes, #{elapsed[2]} seconds."
      puts "Estimated Time Left: #{time_left[0]} hours, #{time_left[1]} minutes, #{time_left[2]} seconds."
      puts "Rate: #{@@avg_rate.to_s[0...5]} queries/minute"
      
      if(@@current_key != 'nominatim')
        puts "Using #{@@current_key} service - #{@@keys[@@current_key][1]} queries left."
      else
        puts "Using #{@@current_key} service - infinite queries left."
      end
      puts "------------------------------------------------------"
    end
  end
  
  def geocode_val(val, count)
    count += 1

    begin
      lat_long = Geocoder.coordinates(val)
      if(!lat_long.nil?)
        @geocoded_vals.push(lat_long)
      else
        @geocoded_vals.push(["", ""])
      end 
      
       @@vals_left -= 1
    rescue Geocoder::OverQueryLimitError   
      if(count < 3)
        sleep(3)
        geocode_val(val, count)
      else
        puts "Error geocoding #{val}. Skipping."
        @geocoded_vals.push(["", ""])
         @@vals_left -= 1
        sleep(3)
      end
    end 
    
    print_info
  end
    
  def geocode_one(key)
    @@current_key = key
    
    configure
    get_const_start
    
    @vals.each do |val|
      geocode_val(val, 0)
    end
    
    return @geocoded_vals
  end
  
  def geocode_mult(key_order)
    key_idx = 3
    @@current_key = key_order[key_idx]
 
    configure
    get_const_start
    count = 0
    
    while(count < @vals.length)
      val = @vals[count]
      if(@@keys[@@current_key][1] > 0 || @@current_key == 'nominatim')
         geocode_val(val, 0)
         count += 1
      elsif(@@keys[@@current_key][1] == 0 && @@current_key != 'nominatim')
        key_idx -= 1
        @@current_key = key_order[key_idx]
        configure
      end
    end

    return @geocoded_vals
  end
  
  def get_const_start
    if(@@const_start_time.nil?)
      @@const_start_time = Time.now
    end
  end
  
  def initialize(vals)
    @vals = []
    @geocoded_vals = []
    @vals = vals
  end
  
  def self.total_vals
    return @@total_vals
  end
  
  def self.set_total_vals(x)
    @@total_vals = x
    @@vals_left = x
  end
  
  def self.keys(x)
    @@keys = x
  end
  
  def self.is_int?(x)
    int_patt = /\A[0-9]+\z/
    
    return x.match(int_patt)
  end
  
  def get_vals
    return @vals
  end
  
  def configure
    Geocoder.configure(
      :lookup => @@current_key.to_sym,
      :api_key => @@keys[@@current_key][0],
      :timeout => 5
    )
  end
end
end



