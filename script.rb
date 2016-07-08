#!/usr/bin/env ruby

require 'mrdialog'
require 'batch_factory'
require 'axlsx'

class Worker
  MIN_DURATION = 10

  def initialize(filename)
    @filename = filename
  end

  def call
    fetch_rows
    p call_stats[102]
  end

  private

  def fetch_rows
    @rows ||= BatchFactory.from_file(@filename, keys: keys).rows.drop(1).select do |row|
      row[:duration] > MIN_DURATION
    end.map do |row|
      row.each_with_object({}) do |(k, v), hash|
        value = v.is_a?(Float) ? v.to_i : v
        hash[k.to_sym] = value
      end
    end
  end

  def calls_by_number
    @calls_by_number ||= @rows.group_by { |row| row[:number] }
  end

  def call_stats
    @call_stats ||= calls_by_number.each_with_object({}) do |(number, calls), hash|
      sorted_calls = calls.sort_by { |call| call[:datetime] }
      hash[number] = {
        first: sorted_calls.first[:datetime],
        last:  sorted_calls.last[:datetime],
        count: sorted_calls.count
      }
    end
  end

  def keys
    [:id, :sip, :datetime, :clid, :number, :state, :duration]
  end

end

class Dialog
  attr_reader :h, :w

  def initialize
    @h = 14
    @w = 48 * 2
  end

  def call
    dialog = MRDialog.new
    dialog.clear = true
    dialog.title = "Please choose a file"

    if file = dialog.fselect(initial_filename, h, w)
      puts '', 'Start'
      Worker.new(file).call
      puts 'Done'
    end
  rescue => e
    puts "#{$!}"
    t = e.backtrace.join("\n\t")
    puts "Error: #{t}"
  end

  private

  def initial_filename
    @initial_filename ||= xlsx_files.sort.last
  end

  def xlsx_files
    @xlsx_files ||= Dir.glob(File.join(File.expand_path('..', __FILE__), '*.xlsx'))
  end
end

Dialog.new.call
