require 'rails_helper'

describe "GeocodePages" do
	describe "Home Page" do 
		it "Should have the content 'Home'" do 
			visit '/geocode_pages/home'
      		expect(page).to have_content('Home')
		end
	end
end
