# -*- encoding : utf-8 -*-
require 'spec_helper'

describe SurfacesExcelExport do
  TEST_FILES     = { without: 'tmp/without_params.xls', decades: 'tmp/for_decades.xls', both: 'tmp/for_both.xls', all: 'tmp/with_all_params.xls' }
  HEADER         = %w( index address side link month1 month2 month3 month4 month5 month6 month7 month8 month9 month10 )
  PART_OF_HEADER = [ { label: '№' }, { label: 'Адрес' }, { label: 'Сторона' }, { label: 'Ссылка на фото, карту' } ]
  AGENDA         = SurfacesExcelExport::STOCK_LEGENDS.values.each.collect {|value| value.sub(value, " - #{value}")}
  UPDATE         = "По состоянию на #{Date.today.strftime('%d.%m.%Y')}"
  HEADER_NAMES   = [ '№', 'Адрес', 'Сторона', 'Ссылка на фото, карту', 'Периоды' ]
  TOP_SHIFT      = 4

  describe '.options_for' do
    it 'for vint' do
      SurfacesExcelExport.options_for(true, false).each do |opt|
        opt[:hide_from_owners].should_not be_true
      end
    end

    it 'for clint' do
      SurfacesExcelExport.options_for(false, true).each do |opt|
        opt[:hide_from_clint].should_not be_true
      end
    end
  end

  it 'prepares formatted price for each country' do
    excel_export = SurfacesExcelExport.new
    file = Tempfile.new('surfaces_excel_export').path + '.xls'
    excel_export.instance_variable_set(:@workbook, WriteExcel.new(file))
    excel_export.send(:prepare_formats)
    Stock::ALL.each do |status|
      COUNTRIES.each do |country, country_hash|
        excel_export.instance_variable_get(:@stock_with_currency_formats)[status][country].should_not be_nil
      end
    end
  end

  describe 'prepares correct excel export' do
    after(:all) do
      delete_test_files
    end

    it 'without params' do
      rows     = { logo: 0, header: 3, first_surface: 5, last_surface: 14, first_agenda: 16, last_agenda: 19 }
      columns  = define_columns(HEADER)
      surfaces = create_surfaces(10, nil, nil)

      generate_tariff(surfaces, MonthTariff)
      prepare_worksheet(surfaces, TEST_FILES[:without], nil)
      check_contacts(columns)

      # Checking value of all the header columns
      check_header(HEADER_NAMES, rows)
      periods_start_column = 4
      check_periods(periods_start_column, false, rows)
      check_value(rows[:logo],         columns[:link],    UPDATE)
      check_value(rows[:first_agenda], columns[:address], MonthTariff::EXPORT_LABELS[0])
      check_agenda(AGENDA, rows)

      # checking 10 surfaces present in the test feed
      check_n_surfaces(columns, 10)
      check_black_header_color(columns, rows)

      h = {month1: :xls_color_47, month2: :xls_color_52, month3: :xls_color_48}
      check_statuses_colors_and_values(h, rows[:first_surface], rows[:last_surface], columns[:month4], columns[:month10], columns[:month1], columns[:month3], columns[:month2], columns)
      check_agenda_colors([:xls_color_48, :xls_color_47, :xls_color_52, :xls_color_53], rows)
    end

    it 'with all possible parameters' do
      rows   = { logo: 0, header: 3, first_surface: 5, last_surface: 7, average: 9 }
      header  = %w(index construction_code surface_code owner_id owner city type format address side grp ots link update month1 month2 month3 month4 month5 month6 month7 month8 month9 month10)
      columns = define_columns(header)
      manager = FactoryGirl.create(:manager)
      surfaces = create_surfaces(3, nil, nil)

      prepare_worksheet(surfaces, TEST_FILES[:all], { enable_our_constructions_numbers: true,
                                                      enable_our_surfaces_numbers:      true,
                                                      enable_owner_codes:               true,
                                                      enable_supertype:                 true,
                                                      enable_types:                     true,
                                                      enable_grp:                       true,
                                                      enable_ots:                       true,
                                                      calculate_average:                true,
                                                      with_owners:                      true,
                                                      city:                             true,
                                                      add_stock_date:                   true,
                                                      hide_all_prices_and_states:       true,
                                                      skip_labels:                      true,
                                                      hide_prices:                      true,
                                                      start_from_current_interval:      true,
                                                      show_prices_for_sold:             true,
                                                      link_to_surfaces_pages:           true,
                                                      create_surfaces_list:             true,
                                                      link_to_surfaces_list:            true,
                                                      manager:                          manager })

      check_header(['№', '№ РК', '№ Пов-ти', '№ Влад-ца', 'Влад-ц', 'Город', 'Тип', 'Формат', 'Адрес', 'Сторона', 'GRP', 'OTS', 'Ссылка на фото, карту', 'Дата обновления', 'Периоды'], rows)

      @sheet.row(rows[:logo]).at(columns[:update]).should   == UPDATE
      @sheet.row(rows[:logo]+1).at(columns[:update]).should == 'Ваша адресная программа на карте'

      periods_start_column = 14
      check_periods(periods_start_column, true, rows)

      # checking 3 surfaces present in the test feed
      first_surface, last_surface = 1, 3
      (first_surface..last_surface).each do |row|
        @sheet.row(row+TOP_SHIFT).at(columns[:index]).to_i.should == row
        row += TOP_SHIFT
        (columns[:construction_code]..columns[:surface_code]).each do |column|
          @sheet.row(row).at(column).should be_kind_of(Float)
        end

        @sheet.row(row).at(columns[:owner]).should match /\AOwner\s*\d+\z/
        @sheet.row(row).at(columns[:city]).should  match /\ARegion\s*\d+\z/
        ['Билборд', '2x4', 'address', 'side X'].each_with_index do |value, index|
          shift = 6
          check_value(row, shift + index, value )
        end
        @sheet.row(row).at(columns[:link]).should include("http://#{PRODUCTION_HOST}/")
      end

      check_black_header_color(columns, rows)

      # checking statuses colors
      (rows[:first_surface]..rows[:last_surface]).each do |row|
        (columns[:month1]..columns[:month10]).each do |column|
          @sheet.row(row).format(column).pattern_fg_color.to_s.should match /(border)|(white)/
        end
      end

    end

    it 'for decades' do
      owner    = FactoryGirl.create(:owner)
      board    = FactoryGirl.create(:construction, owner: owner, type_id: Construction::PERETYAZHKA)
      surfaces = create_surfaces(10, board, nil)

      generate_tariff(surfaces, DecadeTariff)
      prepare_worksheet(surfaces, TEST_FILES[:decades], nil)

      rows    = { logo: 0, header: 3, first_surface: 5, last_surface: 14, first_agenda: 16, last_agenda: 19 }
      header  = %w( index address side link dec1 dec2 dec3 dec4 dec5 dec6 dec7 dec8 dec9 dec10 dec11 dec12 )
      columns = define_columns(header)

      #Checking contacts
      check_contacts(columns)

      # Checking value of all the header columns
      check_header(HEADER_NAMES, rows)
      check_value(rows[:logo], columns[:link], UPDATE)

      (0..3).each do |n|
        check_value(rows[:first_agenda]+n, columns[:address], DecadeTariff::EXPORT_LABELS[n])
      end
      check_value(rows[:last_agenda]+1, columns[:address], DecadeTariff::EXPORT_LINKS[0].last)

      # agenda
      check_agenda(AGENDA, rows)

      # checking 10 surfaces present in the test feed
      check_n_surfaces(columns, 10)
      check_black_header_color(columns, rows)

      h = {dec4: :xls_color_47, dec7: :xls_color_52, dec10: :xls_color_48}
      check_statuses_colors_and_values(h, rows[:first_surface], rows[:last_surface], columns[:dec1], columns[:dec3], columns[:dec4], columns[:dec10], columns[:dec7], columns)

      # checking agenda colors
      check_agenda_colors([:xls_color_48, :xls_color_47, :xls_color_52, :xls_color_53], rows)
    end

    it 'for decades and months' do
      owner = FactoryGirl.create(:owner)
      board = FactoryGirl.create(:construction, owner: owner, type_id: Construction::PERETYAZHKA)
      surfaces_peretyazhki = create_surfaces(10, board, owner)
      surfaces_boards = create_surfaces(5, nil, nil)

      generate_tariff(surfaces_boards,      MonthTariff)
      generate_tariff(surfaces_peretyazhki, DecadeTariff)

      surfaces = surfaces_boards + surfaces_peretyazhki
      prepare_worksheet(surfaces, TEST_FILES[:both], nil)

      rows_months  = { logo: 0, header: 3, first_surface: 5, last_surface:  9, first_agenda: 11, last_agenda: 14 }
      rows_decades = { logo: 0, header:18, first_surface:20, last_surface: 29, first_agenda: 31, last_agenda: 34 }
      header  = HEADER + %w(dec11 dec12)
      columns = define_columns(header)

      #Checking contacts
      check_contacts(columns)

      [rows_months, rows_decades].each do |rows|
        # Checking value of all the header columns
        check_header(HEADER_NAMES, rows)

        # agenda
        check_agenda(AGENDA, rows)
      end
      periods_start_column = 4
      check_periods(periods_start_column, false, rows_months)

      check_value(rows_months[:logo],          columns[:link],    UPDATE)
      check_value(rows_months[:first_agenda],  columns[:address], MonthTariff::EXPORT_LABELS[0])

      (0..3).each do |n|
        check_value(rows_decades[:first_agenda]+n, columns[:address], DecadeTariff::EXPORT_LABELS[n])
      end

      check_n_surfaces(columns, 5)
      check_black_header_color(columns, rows_decades)

      # checking agenda colors
      check_agenda_colors([:xls_color_48, :xls_color_47, :xls_color_52, :xls_color_53], rows_months)
    end
  end

  describe 'all methods behave as intended' do
    before(:each) do
      @manager = FactoryGirl.create(:manager)
      @export = SurfacesExcelExport.new(manager: @manager)
    end

    it 'initialize' do
      @export.class.should                     == SurfacesExcelExport

      variables_hash = { :@columns             => {:d => PART_OF_HEADER, :m => PART_OF_HEADER},
                         :@surfaces            => [],
                         :@columns_widths      => [6, 79, 15, 25],
                         :@logo_join_columns   => 1,
                         :@shift_status_legend => 0,
                         :@discount            => {}
      }
      variables_hash.each do |key, value|
        check_instance_variable(key, value)
      end
    end

    it 'do!' do
      expect{@export.do!}.to raise_error

      surfaces = create_surfaces(10, nil, nil)
      @export.add(surfaces)
      @export.do!.class.should == String
    end

    it 'add' do
      expect{export.add(1)}.to raise_error

      @export.instance_variable_get(:@surfaces).count.should == 0
      surface1, surface2 = FactoryGirl.create(:surface), FactoryGirl.create(:surface)
      @export.add(surface1)
      @export.instance_variable_get(:@surfaces).count.should == 1

      surfaces = [surface1, surface2]
      @export.add(surfaces)
      @export.instance_variable_get(:@surfaces).count.should == 3
    end

    it 'log_export' do
      surface = FactoryGirl.create(:surface)
      @export.add(surface)
      export = @export.log_export
      export.class.should == PG::Result
    end

    it 'prepare_spreadsheet' do
      @export.send(:prepare_spreadsheet)
      @export.filename.should =~ /surfaces_excel_export\d{8}-\d*-\w*\.xls/
      @export.instance_variable_get(:@workbook).class.should  == WriteExcel
      @export.instance_variable_get(:@worksheet).class.should == Writeexcel::Worksheet
    end

    it 'define_columns_order' do
      @export.send(:define_columns_order).should == [:num, :address, :side, :link]
    end
  end

  private

  def check_value(row, column, value)
    @sheet.row(row).at(column).should == value
  end

  def check_agenda(agenda, rows)
    agenda_column = 6
    agenda.each_with_index do |value, index|
      @sheet.row(rows[:first_agenda]+index).at(agenda_column).should == value
    end
  end

  def check_agenda_colors(agenda, rows)
    agenda_colors_column = 5
    agenda.each_with_index do |value, index|
      @sheet.row(rows[:first_agenda]+index).format(agenda_colors_column).pattern_fg_color.should == value
    end
  end

  def check_header(header, rows)
    header.each_with_index do |value, index|
      @sheet.row(rows[:header]).at(index).should  == value
    end
  end

  def check_periods(shift, start_from_current_month, rows)
    periods       = []
    current_year  = Date.today.year
    current_month = Date.today.month

    # adding current month if necessary
    periods << "#{MonthTariff::MONTHS[current_month -1]} #{current_year}" if start_from_current_month

    # getting all months that left in this year
    (current_month..11).each do |month|
      periods << "#{MonthTariff::MONTHS[month]} #{current_year}"
    end

    # getting months for the next year until the current month
    (0..current_month).each do |month|
      periods << "#{MonthTariff::MONTHS[month]} #{current_year.next}"
    end

    # Checking the months intervals in the header
    periods[0..9].each_with_index do |period, index|
      @sheet.row(rows[:header]+1).at(index+shift).should == period
    end
  end

  def define_columns(header)
    columns = {}
    header.each_with_index do |value, index|
      columns[value.to_sym] = index
    end
    columns
  end

  def delete_test_files
    TEST_FILES.each_pair do |_, file|
      File.delete(file) if File.exists?(file)
    end
  end

  def create_surfaces(quantity, board, owner)
    owner ||= FactoryGirl.create(:owner)
    board ||= FactoryGirl.create(:construction, owner: owner, type_id: Construction::BILBOARD_2X4)
    surfaces = []
    quantity.times do
      surfaces << FactoryGirl.create(:surface, construction: board)
    end
    surfaces
  end

  def check_n_surfaces(columns, n)
    first_surface, last_surface = 1, n
    (first_surface..last_surface).each do |row|
      @sheet.row(row+TOP_SHIFT).at(columns[:index]).to_i.should == row
      row += TOP_SHIFT
      check_value(row, columns[:address], 'address')
      check_value(row, columns[:side],    'side X')
      @sheet.row(row).at(columns[:link]).should include("http://#{PRODUCTION_HOST}/#")
    end
  end

  def generate_tariff(surfaces, tariff)
    today  = Time.zone.today
    # hash for generating some statuses and tariffs.
    # FREE for next month, SOLD for +2 months, RESERVED for +3 months
    generate_tariffs = { 1 => Stock::FREE, 2 => Stock::SOLD, 3 => Stock::RESERVED }
    generate_tariffs.each_pair do |key, value|
      surfaces.each do |surface|
        status = {status: value, base: 111, sell: 222, our: 111}
        surface.set_stock(tariff.new(today + key.month),      status) if tariff == MonthTariff
        surface.set_stock(tariff.new(today + key.month).next, status) if tariff == DecadeTariff
      end
    end
  end

  def check_contacts(columns)
    [SurfacesExcelExport::EMAIL, SurfacesExcelExport::PHONE].each do |value|
      @sheet.row(0).at(columns[:index]).should include(value)
    end
  end

  def check_black_header_color(columns, rows)
    columns.each_with_index do |_, index|
      @sheet.row(rows[:header]).format(index).pattern_fg_color.should == :xls_color_54
    end
  end

  def check_statuses_colors_and_values(statuses, start_row, end_row, start_column1, end_column1, start_column2, end_column2, check_value_column, columns)
    (start_row..end_row).each do |row|
      (start_column1..end_column1).each do |column|
        @sheet.row(row).format(column).pattern_fg_color.should == :xls_color_53
      end
      [start_column2, end_column2].each do |column|
        check_value(row, column, 222.0)
      end
      check_value(row, check_value_column, 'Продано')

      statuses.each_pair do |key, value|
        @sheet.row(row).format(columns[key.to_sym]).pattern_fg_color.should == value
      end
    end
  end

  def prepare_worksheet(surfaces, filename, params)
    params ||= {}
    export = SurfacesExcelExport.new(params)
    export.add(surfaces)
    book = Spreadsheet.open(export.do!)
    book.write(filename)
    @sheet = book.worksheet 0
  end

  def check_instance_variable(variable, value)
    @export.instance_variable_get(variable).should == value
  end
end 
