require 'GeocodeController.rb'

at_exit do 
  last = GeocodeController.new
  last.write_keys
end

class GeocodePagesController < ApplicationController	
	@books_in = []
	
	def home
	end
	
	def geocode
		books_in = Spreadsheet.open(params[:file].tempfile)
		row = params[:row][0].to_i
		col = params[:col][0].to_i
		filename = params[:file].original_filename
		test = GeocodeController.new(row, col, books_in, filename)
		test.geocode
		
		send_file "#{filename.split('.')[0]}_geocoded.xls", :type=>"application/vnd.ms-excel", :x_sendfile=>true
	end
end
