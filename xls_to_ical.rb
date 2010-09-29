require "rubygems"
require "icalendar"
require "roo"

C_DATE = "A"
C_START_TIME = "C"
C_END_TIME = "D"
C_NAME = "E"
C_TYPE = "F"
C_ROOM = "H"
C_GRUPA_DZIEKANACKA = "J"
C_GRUPA_LEKTORSKA = "K"
C_GRUPA_WYKLADOWA = "L"
C_GAME_MASTER = "M"

def combine_seconds_with_date(date, seconds)
	hour = seconds / 60 / 60
	minutes = seconds / 60.0 % 60.0
	
	DateTime.civil(date.year, date.month, date.day, hour.to_i, minutes.to_i)
end

def main
	file_name = ARGV.first

	puts "Reading file #{file_name}"
	@cal = Icalendar::Calendar.new

	xl = Excel.new(file_name)
	xl.default_sheet = xl.sheets.first
	
	4.upto(xl.last_row) do |row|
		begin
			event = Icalendar::Event.new
			event.summary = xl.cell(row, C_NAME)
			
			date = xl.cell(row, C_DATE)
			next if date.nil?
			
			event.start = combine_seconds_with_date(date,xl.cell(row, C_START_TIME))
			event.end = combine_seconds_with_date(date,xl.cell(row, C_END_TIME))
			event.location = xl.cell(row, C_ROOM)
			event.organizer = xl.cell(row, C_GAME_MASTER)
			
			description = "Forma zajęć: #{xl.cell(row, C_ROOM)}\n"
			description += "Grupa dziekańska: #{xl.cell(row, C_GRUPA_DZIEKANACKA)}\n"
			description += "Grupa lektorska/komputerowa: #{xl.cell(row, C_GRUPA_LEKTORSKA)}\n"
			description += "Grupa wykładowa: #{xl.cell(row, C_GRUPA_WYKLADOWA)}\n"
			description += "Wykładowca: #{xl.cell(row, C_GAME_MASTER)}\n"
			description += "Sala: #{xl.cell(row, C_ROOM)}"
			
			event.description = description
			puts "#{row}: Date: #{xl.cell(row, C_DATE)} #{xl.cell(row, C_START_TIME)} -> #{xl.cell(row, C_END_TIME)} >> #{xl.cell(row, C_NAME)}"
			@cal.add_event(event)
		rescue Exception => e
			puts "Could not parse row #{row}: #{e.to_s}"
			puts e.backtrace.join("\n")
		end
	end
	
	file = File.new("lessons.ics", "w")
	file.write @cal.to_ical
	file.close

end

main

#

#event.dtstart = start_time
#event.dtend = end_time
#event.summary = name
#event.description = description
#