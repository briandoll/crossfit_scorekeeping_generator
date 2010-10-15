require 'writeexcel'

module Crossfit
  module Scorekeeping
    class Generator
      
      attr_accessor :worksheet, :bold_format, :final_score_column, :final_rank_column, :athlete_rows, :rank_columns
      
      def initialize(number_of_athletes = 25, number_of_events = 5)
        file_name = "#{Date.today.strftime('%Y_%m_%d_crossfit_scores.xls')}"
        workbook = WriteExcel.new(file_name)
        @worksheet = workbook.add_worksheet
        @bold_format = workbook.add_format
        @bold_format.set_bold
        generate(number_of_athletes, number_of_events)
        workbook.close
      end

      def generate(number_of_athletes, number_of_events)
        @final_score_column = num_to_alpha[((number_of_events * 2) + 3)]
        @final_rank_column = num_to_alpha[((number_of_events * 2) + 4)]
        @rank_columns = []
        @athlete_rows = [4, (number_of_athletes + 3)]
        write_competition_headers(number_of_events)
        write_athlete_rows(number_of_athletes, number_of_events)
      end
      
      def write_competition_headers(number_of_events)
        @worksheet.write("A1","Athlete Name", @bold_format)
        (1..number_of_events).each do |position|
          # account for initial column, plus two colums per previous event
          prefix = (1.eql?(position)) ? 1 : ((position - 1) * 2 + 1)
          event_number = position
          score_letter = num_to_alpha[prefix]
          rank_letter  = num_to_alpha[prefix + 1]
          @rank_columns << rank_letter
          @worksheet.write("#{score_letter}1", "Event ##{event_number} Score", @bold_format)
          @worksheet.write("#{rank_letter}1",  "Event ##{event_number} Rank",  @bold_format)
          @worksheet.write("#{rank_letter}2", 0, @bold_format)
        end
        @worksheet.write("#{@final_score_column}1","Final Score", @bold_format)
        @worksheet.write("#{@final_rank_column}1","Final Rank", @bold_format)
      end
      
      def write_athlete_rows(number_of_athletes, number_of_events)
        (4..(number_of_athletes + 3)).each do |row_num|
          @worksheet.write("A#{row_num}","Athlete Name")
          last_used_column = 0
          (1..number_of_events).each do |event_num|
            score_column = last_used_column + 1
            rank_column  = last_used_column + 2

            @worksheet.write("#{num_to_alpha[score_column]}#{row_num}", 0)  # score
            event_range = "#{num_to_alpha[score_column]}#{athlete_rows.first}:#{num_to_alpha[score_column]}#{athlete_rows.last}"
            @worksheet.write("#{num_to_alpha[rank_column]}#{row_num}", "=RANK(#{num_to_alpha[score_column]}#{row_num},#{event_range},#{num_to_alpha[rank_column]}2)") # rank

            @worksheet.write("#{final_score_column}#{row_num}", "=SUM(#{athlete_rank_rows(row_num)})") # sum of ranks
            
            final_score_range = "#{final_score_column}#{athlete_rows.first}:#{final_score_column}#{athlete_rows.last}"
            @worksheet.write("#{final_rank_column}#{row_num}", "=RANK(#{final_score_column}#{row_num}, #{final_score_range}, 1)") # rank of ranks
            last_used_column += 2
          end
        end
      end
      
      def athlete_rank_rows(row_num)
        rank_columns.map{|column| "#{column}#{row_num}"}.flatten.join(",")
      end
      
      # heh... funny i've never had to write this before...
      def num_to_alpha
        @num_to_alpha ||= begin
          numbers = (0..25).to_a
          alpha   = ('A'..'Z').to_a
          mapping = {}
          numbers.each do |n|
            mapping[n] = alpha[n]
          end
          mapping
        end
      end
      
    end
  end
end

Crossfit::Scorekeeping::Generator.new 25, 6