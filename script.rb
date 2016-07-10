#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

require 'mrdialog'
require 'batch_factory'
require 'axlsx'

class Worker
  attr_reader :viewer

  class Parser
    attr_reader :rows, :filename
    MIN_DURATION = 10

    def initialize(filename)
      @filename = filename
    end

    def call_stats(type)
      method_name = "#{type}_calls_by_number".to_sym
      stats = send(method_name).each_with_object({}) do |(number, calls), hash|
        sorted_calls = calls.sort_by { |call| call[:datetime] }
        hash[number] = {
          main: {
            first: sorted_calls.first[:datetime],
            last:  sorted_calls.last[:datetime],
            count: sorted_calls.count
          }
        }
      end

      special_numbers.each do |number|
        before_17 = send(method_name)[number].select { |call| call[:datetime].hour < 17 }.sort_by { |call| call[:datetime] }
        after_17  = send(method_name)[number].select { |call| call[:datetime].hour >= 17 }.sort_by { |call| call[:datetime] }

        stats[number] = {
          before_17: {
            first: before_17.first ? before_17.first[:datetime] : nil,
            last:  before_17.last ? before_17.last[:datetime] : nil,
            count: before_17.count
          },
          after_17: {
            first: after_17.first ? after_17.first[:datetime] : nil,
            last:  after_17.last ? after_17.last[:datetime] : nil,
            count: after_17.count
          }
        }
      end

      stats
    end

    private

    def rows
      @rows ||= BatchFactory.from_file(@filename, keys: keys).rows.drop(1).select do |row|
        row[:duration] > MIN_DURATION
      end.map do |row|
        row.each_with_object({}) do |(k, v), hash|
          value = v.is_a?(Float) ? v.to_i : v
          hash[k.to_sym] = value
        end
      end
    end

    def incoming_calls
      @incoming_calls ||= rows.select { |call| inner_number?(call[:number]) || call[:number] == 'incoming' }
    end

    def incoming_calls_by_number
      @incoming_calls_by_number ||= incoming_calls.group_by { |call| call[:number] }
    end

    def outgoing_calls
      @outgoing_calls ||= rows.select { |call| inner_number? call[:sip] }
    end

    def outgoing_calls_by_number
      @outgoing_calls_by_number ||= outgoing_calls.group_by { |call| call[:sip] }
    end

    def keys
      [:id, :sip, :datetime, :clid, :number, :state, :duration]
    end

    def special_numbers
      [110, 111]
    end

    def inner_number?(number)
      (100...400).include? number
    end
  end

  class Viewer
    attr_reader :parser, :styles

    def initialize(parser)
      @parser = parser
    end

    def call
      Axlsx::Package.new do |p|
        workbook = p.workbook
        add_sheet workbook, { name: 'Incoming', type: :incoming }
        add_sheet workbook, { name: 'Outgoing', type: :outgoing }
        p.serialize 'results.xlsx'
      end
    end

    private

    def define_styles(sheet)
      @styles ||= begin
        centered = sheet.styles.add_style alignment: { horizontal: :center }
        header   = sheet.styles.add_style alignment: { horizontal: :center }, bg_color: 'FFE699', border: { style: :thin, color: '000000', edges: [:top, :bottom, :left, :right] }
        number   = sheet.styles.add_style bg_color: 'DDEBF7', border: { style: :thin, color: '000000', edges: [:top, :bottom, :left, :right] }
        details  = sheet.styles.add_style bg_color: 'BDD7EE', border: { style: :thin, color: '000000', edges: [:top, :bottom, :left, :right] }

        { centered: centered, header: header, phone_number: number, details: details }
      end
    end

    def add_headers(sheet)
      sheet.merge_cells 'C1:E1'
      sheet.add_row [nil, nil, date], style: [nil, nil, styles[:header]]
      # sheet.add_row [nil, nil, 'Week day'], style: styles[:header]
      sheet.add_row [nil, nil, 'Кол-во', 'Первый зв.', 'Последний зв.'], widths: [:auto, :auto, 20, 20, 20], style: [nil, nil, styles[:header], styles[:header], styles[:header]]

      # sheet.merge_cells 'C1:E1'
      # sheet.merge_cells 'C2:E2'
    end

    def add_sheet(book, params)
      book.add_worksheet(name: params[:name]) do |sheet|
        define_styles sheet
        add_headers sheet

        Hash[parser.call_stats(params[:type]).sort_by { |k, _| k.to_i }].each do |number, stats|
          if stats.has_key? :main
            sheet.add_row [
              number,
              nil,
              stats[:main][:count],
              stats[:main][:first]&.strftime('%H:%M:%S'),
              stats[:main][:last]&.strftime('%H:%M:%S'),
            ], style: [styles[:phone_number], styles[:details], styles[:centered], styles[:centered], styles[:centered]]
          else
            sheet.add_row [number, nil], style: [styles[:phone_number], styles[:details], nil, nil, nil]
            sheet.add_row [
              nil,
              'До 17:00',
              stats[:before_17][:count],
              stats[:before_17][:first]&.strftime('%H:%M:%S'),
              stats[:before_17][:last]&.strftime('%H:%M:%S'),
            ], style: [styles[:phone_number], styles[:details], styles[:centered], styles[:centered], styles[:centered]]
            sheet.add_row [
              nil,
              'После 17:00',
              stats[:after_17][:count],
              stats[:after_17][:first]&.strftime('%H:%M:%S'),
              stats[:after_17][:last]&.strftime('%H:%M:%S'),
            ], style: [styles[:phone_number], styles[:details], styles[:centered], styles[:centered], styles[:centered]]
          end
        end

        total = parser.call_stats(params[:type]).inject(0) do |count, (number, stats)|
          if stats.has_key? :main
            count + stats[:main][:count]
          else
            count + stats[:before_17][:count] + stats[:after_17][:count]
          end
        end
        sheet.add_row [
          nil,
          'Сумма',
          total
        ], style: [styles[:phone_number], styles[:details], styles[:centered]]
      end
    end

    def sort
      -> (a, b) do
        return -1 if a.is_a? String
        a <=> b
      end
    end

    def date
      File.basename(parser.filename).split('.').first(3).join '/'
    end
  end

  def initialize(filename)
    @parser = Parser.new(filename)
    @viewer = Viewer.new(@parser)
  end

  def call
    viewer.call
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
    dialog.title = 'Please choose a file'

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
