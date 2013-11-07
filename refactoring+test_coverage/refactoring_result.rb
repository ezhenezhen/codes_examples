# -*- encoding : utf-8 -*-
class SurfacesExcelExport
  include ActionView::Helpers::NumberHelper
  include ConstructionsHelper

  attr_reader :filename

  STOCK_LEGENDS = {
      Stock::RESERVED => 'Чужой резерв',
      Stock::FREE     => 'Свободно',
      Stock::SOLD     => 'Продано',
      Stock::UNKNOWN  => 'Необходимо уточнить наличие'
  }

  OPTIONS = [
      {                                       description: 'Добавить колонки:'}, # group header
      {id: :enable_our_constructions_numbers, description: 'Наш № РК',                         hide_from_owners: true                                       },
      {id: :enable_our_surfaces_numbers,      description: 'Наш № поверхности',                hide_from_owners: true,                        default: true },
      {id: :enable_owner_codes,               description: '№ поверхности владельца',                                  hide_from_clint: true,               },
      {id: :enable_supertype,                 description: 'Тип'                                                                                            },
      {id: :enable_types,                     description: 'Формат'                                                                                         },
      {id: :enable_grp,                       description: 'GRP'                                                                                            },
      {id: :enable_ots,                       description: 'OTS'                                                                                            },
      {id: :calculate_average,                description: 'Средние показатели'                                                                             },
      {id: :with_owners,                      description: 'Владелец',                         hide_from_owners: true, hide_from_clint: true                },
      {id: :city,                             description: 'Город'                                                                                          },
      {id: :add_stock_date,                   description: 'Дата обновления статуса'                                                                        },
      {id: :hide_all_prices_and_states,       description: 'Скрыть все статусы и цены'                                                                      },
      {id: :skip_labels,                      description: 'Пропустить комментарии'                                                                         },
      {                                       description: 'Настройки:',                       hide_from_owners: true                                       },
      {id: :hide_prices,                      description: 'Режим лояльности (скрывать цены)', hide_from_owners: true, hide_from_clint: true                },
      {id: :create_surfaces_list,             description: 'Создать страницу АП',              hide_from_owners: true                                       },
      {id: :start_from_current_interval,      description: 'Начать с текущего месяца/декады',  hide_from_owners: true,                       default: lambda{Time.now.day <= 5}},
      {id: :show_prices_for_sold,             description: 'Показывать цены для проданных'                                                                  },
      {                                       description: 'Ставить ссылку на:',               hide_from_owners: true                                       },
      {id: :link_to_surfaces_list,            description: 'Страницу АП',                      hide_from_owners: true                                       },
      {id: :link_to_surfaces_pages,           description: 'Страницу стороны',                 hide_from_owners: true                                       },
  ]

  EMAIL = 'xxx'
  PHONE = 'xxx'

  HEADER_ROWS = 3

  def initialize(opts = {})
    @options  = opts
    @surfaces = []
    @discount = {}

    define_columns
    deserialize_discounts

    @protocol_host_port ||= "http://#{PRODUCTION_HOST}"
  end

  def do!
    raise 'No surfaces to export' if @surfaces.blank?
    prepare_spreadsheet
    @surfaces_list    = SurfacesList.create(manager: @options[:manager]) if (@options[:create_surfaces_list] || @options[:link_to_surfaces_list])
    @surfaces_by_type = Surface.separate_by_tariff(@surfaces)

    estimate_rows
    set_columns_width
    prepare_formats
    insert_logo
    export_information

    @workbook.close
    log_export

    @filename
  end

  def self.options_for(vint, clint)
    opts = []
    OPTIONS.each do |o|
      next if o[:hide_from_owners] && vint
      next if o[:hide_from_clint]  && clint
      opts << o
    end
    opts
  end

  # Add surface or list of surfaces to export
  def add(surface)
    case surface
      when Surface then @surfaces << surface
      when Array then surface.each{|s| add(s)}
      else raise 'Can add only surfaces or array of surfaces'
    end
  end

  # Log this export to DB
  def log_export
    sql = <<-SQL
      INSERT INTO surfaces_excel_exports(manager_id, options, surfaces_ids, created_at) VALUES (#{@options[:manager].is_a?(Manager) ? @options[:manager].id : 'NULL'}, '#{@options.to_json}', '#{@surfaces.map(&:id) * ','}', now())
    SQL
    ActiveRecord::Base.connection.execute(sql)
  end

  def surfaces_list_link
    @surfaces_list.link
  end

  def self.clear_parameters(params, manager)
    h = {}
    OPTIONS.each do |opt|
      next if params["excel_option_#{opt[:id]}"].blank?
      next if opt[:hide_from_owners] && manager.department.owner # Don't pass hidden from owners parameters if current manager is from owner
      next if opt[:hide_from_clint]  && manager.clint?           # Don't pass hidden from clint managers parameters if current manager is clint
      h[opt[:id]] = true
    end
    h
  end

  private

  def deserialize_discounts
    @options[:discounts].split(',').each { |d| dd = d.split(':'); @discount[dd.first.to_i] = dd.last.to_i } if @options[:discounts]
  end

  def define_columns
    @columns = {}
    @columns_widths = []
    @logo_join_columns = 1
    @shift_status_legend = 0

    define_for_decades
    define_for_months
  end

  def prepare_spreadsheet
    @filename = Tempfile.new('surfaces_excel_export').path + '.xls'
    @workbook = WriteExcel.new(@filename)
    @worksheet = @workbook.add_worksheet(Date.today.strftime('%d.%m.%Y'))
    @worksheet.set_zoom(75)
  end

  def estimate_rows
    estimate_rows =  @surfaces.length
    estimate_rows += (@surfaces_by_type.length - 1) * 2 # two rows between types
    estimate_rows += @surfaces_by_type.length           # row between table and labels
    estimate_rows += @surfaces_by_type.length * 2       # 2 rows for each headers

    @surfaces_by_type.keys.each do |p|
      t = Tariff.class_by_prefix(p)
      estimate_rows += [t::EXPORT_LABELS.length + t::EXPORT_LINKS.length, 4].max + 1

      # onw row spacer and one row for average line
      estimate_rows += 2 if @options[:calculate_average]
    end
    (0..(HEADER_ROWS - 1)).each{|i| @worksheet.set_row(i, 25)}
    (HEADER_ROWS..(HEADER_ROWS + estimate_rows)).each{|i| @worksheet.set_row(i, 21)}
  end

  def insert_logo
    @worksheet.merge_range(0, 0, 1, @logo_join_columns, "Телефон: #{@options[:phone] || PHONE}\nEmail: #{@options[:email] || EMAIL}", @logo_format)

    logo = (@options[:logo] || File.join(Rails.root, %w(public i logo.png)))
    @worksheet.insert_image('A1', logo, 10, 10, 1, 1.1)
    @worksheet.write(0, @logo_join_columns + 2, "По состоянию на #{Date.today.strftime('%d.%m.%Y')}", @timestamp_format)
  end

  def export_information
    row = HEADER_ROWS
    @surfaces_by_type.each do |tariff_prefix, surfaces|

      tariff = Tariff.class_by_prefix(tariff_prefix)
      intervals = []

      # Start from current or next interval
      detect_start_interval(intervals, tariff)

      # Writing headers and joining cells in headers
      row = export_interval_headers(intervals, row, tariff_prefix)

      # Define columns order
      order = define_columns_order

      dataset = []
      prices_dataset = {}

      row = export_surfaces_information(dataset, intervals, order, prices_dataset, row, surfaces)

      # Writing average values if required
      row = calculate_average_values(dataset, intervals, order, prices_dataset, row)

      # Writing labels and links under the table
      row = write_statuses_legend(row, tariff)

      # Adding link to Active Sales page
      add_link_to_active_sales_page

      row += 3
    end
  end

  def export_surfaces_information(dataset, intervals, order, prices_dataset, row, surfaces)
    surfaces.each_with_index do |surface, index|
      @surfaces_list.add(surface) if @options[:create_surfaces_list] || @options[:link_to_surfaces_list]
      row += 1

      num = index + 1
      data = prepare_data(num, surface)
      dataset << data

      order.each_with_index do |id, idx|
        #f = @cell_formats[id][num.odd? ? :odd : :even]
        f = @cell_formats[id][:even]
        if id == :link
          @worksheet.write_url(row, idx, data[id], data[id], f)
        else
          @worksheet.write(row, idx, data[id], f)
        end
      end

      export_prices_and_states(intervals, prices_dataset, row, surface)
    end
    row
  end

  def export_interval_headers(intervals, row, tariff_prefix)
    @columns[tariff_prefix].each_with_index do |column, index|
      @worksheet.merge_range(row, index, row + 1, index, column[:label], @header_cell_format)
    end
    @worksheet.merge_range(row, @total_columns, row, @total_columns + intervals.length - 1, 'Периоды', @header_cell_format)
    row += 1

    # Writing interval headers
    intervals.each_with_index do |interval, i|
      @worksheet.write(row, @total_columns + i, interval.text_name, @header_cell_format_intervals)

      # Set width for interval columns
      @worksheet.set_column(@total_columns + i, @total_columns + i, 15)
    end
    row
  end

  def detect_start_interval(intervals, tariff)
    shift_number = (@options[:start_from_current_interval] ? 0 : 1)
    tariff.show_intervals.times { |i| intervals << tariff.now.next(i + shift_number) }
  end

  def set_columns_width
    @total_columns = @columns_widths.length
    @columns_widths.each_with_index do |width, index|
      @worksheet.set_column(index, index, width)
    end

    max_intervals_shown = @surfaces_by_type.keys.map { |k| Tariff.class_by_prefix(k).show_intervals }.max
    @worksheet.set_column(@total_columns, max_intervals_shown + @total_columns - 1, 20)
  end

  def prepare_data(num, surface)
    {
        num:              num,
        construction_id:  surface.construction.id,
        surface_id:       surface.id,
        owner_code:       surface.owner_code,
        owner:            surface.construction.owner.name,
        city:             surface.construction.region.name,
        supertype:        layout_supertype_singular(surface.construction.supertype),
        type:             surface.construction.type_name,
        grp:              surface.grp,
        ots:              surface.ots,
        stock_updated_at: (dt = surface.stock_updated_at; dt ? dt.strftime('%d.%m.%Y') : ''),
        address:          layout_address(surface),
        side:             layout_side(surface),
        link:             link(surface)
    }
  end

  def add_link_to_active_sales_page
    @worksheet.write_url(1, @logo_join_columns + 2, @surfaces_list.link(@protocol_host_port), 'Ваша адресная программа на карте', @link_format) if @options[:create_surfaces_list]
  end

  def write_statuses_legend(row, tariff)
    unless @options[:skip_labels]
      row += 2
      first_label_row = row
      (tariff::EXPORT_LABELS + tariff::EXPORT_LINKS).each do |label|
        if label.is_a?(Array)
          @worksheet.write_url(row, 1, label.first, label.second, @link_format)
        else
          @worksheet.write(row, 1, label, @timestamp_format)
        end
        row += 1
      end

      # Writing statuses legend
      row = first_label_row # + 2
      [Stock::RESERVED, Stock::FREE, Stock::SOLD, Stock::UNKNOWN].each do |status|
        @worksheet.write(row, @shift_status_legend + 5, '', @stock_formats[status])
        @worksheet.write(row, @shift_status_legend + 6, ' - ' + STOCK_LEGENDS[status], @stock_legend_format)
        row += 1
      end
    end
    row
  end

  def export_prices_and_states(intervals, prices_dataset, row, surface)
    unless @options[:hide_all_prices_and_states]
      col    = @total_columns
      prices = surface.all_prices_hash

      intervals.each do |interval|
        prices_dataset[interval] ||= []
        state = surface.stock_at(interval)

        if state == Stock::SOLD && !(@options[:show_prices_for_sold])
          value = 'Продано'
        else
          value = (@options[:hide_prices].blank? ? prepare_price_value(prices, interval, surface.construction.owner_id) : '')
          prices_dataset[interval] << value if value.present?
        end
        @worksheet.write(row, col, value, @stock_formats[state])
        col += 1
      end
    end
  end

  def define_columns_order
    order = []
    order << :num
    order << :construction_id  if @options[:enable_our_constructions_numbers]
    order << :surface_id       if @options[:enable_our_surfaces_numbers]
    order << :owner_code       if @options[:enable_owner_codes]
    order << :owner            if @options[:with_owners]
    order << :city             if @options[:city]
    order << :supertype        if @options[:enable_supertype]
    order << :type             if @options[:enable_types]
    order << :address
    order << :side
    order << :grp              if @options[:enable_grp]
    order << :ots              if @options[:enable_ots]
    order << :link
    order << :stock_updated_at if @options[:add_stock_date]
    order
  end

  def calculate_average_values(dataset, intervals, order, prices_dataset, row)
    if @options[:calculate_average]
      row += 2
      # Label and average mediaindicators
      @worksheet.write(row, order.index(:address), 'Средние показатели', @average_formats[:label])
      @worksheet.write(row, order.index(:grp),     '%0.2f' % dataset.map { |d| d[:grp] }.compact.avg, @average_formats[:grp]) if @options[:enable_grp]
      @worksheet.write(row, order.index(:ots),     '%0.2f' % dataset.map { |d| d[:ots] }.compact.avg, @average_formats[:ots]) if @options[:enable_ots]

      #Calculating and writing percentage of A sides
      calculate_a_sides(dataset, order, row)

      # Writing average prices for each month
      col = @total_columns
      intervals.each do |interval|
        next unless prices_dataset[interval]
        avg = prices_dataset[interval].avg
        @worksheet.write(row, col, avg > 0 ? avg.to_i : '', @stock_formats[Stock::UNKNOWN])
        col += 1
      end
    end
    row
  end

  def calculate_a_sides(dataset, order, row)
    sides = dataset.map { |d| d[:side] }.compact
    a_sides = sides.select { |s| Surface.is_side_a?(s) }
    a_sides_percentage = ((a_sides.length.to_f / sides.length.to_f) * 100).to_i
    @worksheet.write(row, order.index(:side), "Сторон А: #{a_sides_percentage}%", @average_formats[:side])
  end

  def define_for_months
    @columns[:m] ||= []
    columns = [prepare_months_columns('№',                     true),
               prepare_months_columns('№ РК',                  @options[:enable_our_constructions_numbers]),
               prepare_months_columns('№ Пов-ти',              @options[:enable_our_surfaces_numbers]),
               prepare_months_columns('№ Влад-ца',             @options[:enable_owner_codes]),
               prepare_months_columns('Влад-ц',                @options[:with_owners]),
               prepare_months_columns('Город',                 @options[:city]),
               prepare_months_columns('Тип',                   @options[:enable_supertype]),
               prepare_months_columns('Формат',                @options[:enable_types]),
               prepare_months_columns('Адрес',                 true),
               prepare_months_columns('Сторона',               true),
               prepare_months_columns('GRP',                   @options[:enable_grp]),
               prepare_months_columns('OTS',                   @options[:enable_ots]),
               prepare_months_columns('Ссылка на фото, карту', true),
               prepare_months_columns('Дата обновления',       @options[:add_stock_date])
    ]
    columns.each do |hash|
      add_column(:m, label: hash[:label]) if hash[:condition]
    end
  end

  def prepare_months_columns(label, condition)
    { label: label, condition: condition }
  end

  def adjust_export_view(columns_widths, logo_join_columns, shift_status_legend)
    @columns_widths      << columns_widths
    @logo_join_columns   += logo_join_columns
    @shift_status_legend += shift_status_legend
  end

  def add_column(interval_type, column)
    @columns[interval_type] << column
  end

  def define_for_decades
    @columns[:d] ||= []
    columns = [
        prepare_decades_columns('№',                     [6,  0, 0], true),
        prepare_decades_columns('№ РК',                  [16, 1, 0], @options[:enable_our_constructions_numbers]),
        prepare_decades_columns('№ Пов-ти',              [16, 1, 0], @options[:enable_our_surfaces_numbers]),
        prepare_decades_columns('№ Влад-ца',             [16, 1, 0], @options[:enable_owner_codes]),
        prepare_decades_columns('Влад-ц',                [16, 1, 1], @options[:with_owners]),
        prepare_decades_columns('Город',                 [16, 1, 1], @options[:city]),
        prepare_decades_columns('Тип',                   [16, 1, 1], @options[:enable_supertype]),
        prepare_decades_columns('Формат',                [16, 1, 1], @options[:enable_types]),
        prepare_decades_columns('Адрес',                 [79, 0, 0], true),
        prepare_decades_columns('Сторона',               [15, 0, 0], true),
        prepare_decades_columns('GRP',                   [10, 1, 1], @options[:enable_grp]),
        prepare_decades_columns('OTS',                   [10, 1, 1], @options[:enable_ots]),
        prepare_decades_columns('Ссылка на фото, карту', [25, 0, 0], true),
        prepare_decades_columns('Дата обновления',       [16, 1, 1], @options[:add_stock_date])
    ]
    columns.each do |hash|
      if hash[:condition]
        add_column(:d, label: hash[:label])
        adjust_export_view(hash[:adjust][0], hash[:adjust][1], hash[:adjust][2])
      end
    end
  end

  def prepare_decades_columns(label, adjust, condition)
    {label: label, adjust: adjust, condition: condition}
  end

  def prepare_price_value(prices, interval, owner_id)
    p = prices[interval.id]
    return '' if p.blank?
    price = p[:sell] || p[:base]

    # Enable discount if present for this owner
    if price && @discount[owner_id]
      delta = (@options[:discount_unit]      == '%' ? price * @discount[owner_id] / 100 : @discount[owner_id])
      price = (@options[:discount_direction] == '-' ? price - delta : price + delta)
    end

    price
  end

  def layout_address(surface)
    surface.construction.address
  end

  def layout_side(surface)
    surface.name.gsub('Сторона ','')
  end

  def price(val)
    val == 0 ? ' ' : number_to_currency(val, unit: '', precision: 0, delimiter: ' ', format: '%n')
  end

  # Prepare link for given surface depend on parameters
  def link(s)
    if    @options[:link_to_external_site]
      s.exteral_site_link(@protocol_host_port)
    elsif @options[:hide_prices]
      # Links to surface page without prices. For loyal clients
      s.link_to_noprice_page(@protocol_host_port)
    elsif @options[:link_to_surfaces_list]
      # Links to SurfacesList page with anchor to exact surface
      @surfaces_list.link_for_surface(s, @protocol_host_port)
    elsif @options[:link_to_surfaces_pages]
      # Links to surface page with all details (with price)
      s.link(@protocol_host_port)
    else
      # By default insert links to index map with construction details in side panel
      s.construction.link_to_page(@protocol_host_port)
    end
  end

  def prepare_formats
    set_default_cell_props
    set_stock_formats
    set_custom_colors
    set_header_cell_format
    set_logo_format
    set_timestamp_format
    set_links_formats
    set_stock_legend_format
    set_cell_formats_props(header_format, prices_format)
    set_cell_formats
    set_average_formats(average_format)
  end

  def set_default_cell_props
    @default_cell_props = {border: 1, valign: 'vcenter', font:   'Arial', size:   11}
  end

  def set_stock_legend_format
    @stock_legend_format = @workbook.add_format(merge_formats({align: 'left', bg_color: 'white', border_color: 'black'}))
  end

  def merge_formats(hash)
    @default_cell_props.merge(hash)
  end

  def set_header_cell_format
    @header_cell_format = @workbook.add_format(header_cell_format_props)
    # Create totally copy of usual styles. Don't know why but Excel from Office XP
    # crash on exports generated with the same styles for intervals headers cells
    # So that is just workaround for Excel XP
    @header_cell_format_intervals = @workbook.add_format(header_cell_format_props)
  end

  def header_cell_format_props
    {border: 1, valign: 'vcenter', align: 'center', bg_color: @header_gray, color: 'white', bold: 700, size: 11, text_wrap: 1}
  end

  def add_format(hash)
    @workbook.add_format(merge_formats(hash))
  end

  def set_links_formats
    @labels_link_format = add_format({align: 'left', size: 13, color: 'blue', border_color: 'gray'})
    @link_format        = add_format({border: 0, align: 'left', size: 13, color: 'blue', underline: 1})
  end

  def set_timestamp_format
    @timestamp_format   = add_format({border: 0, size: 13, align: 'left'})
  end

  def set_logo_format
    @logo_format        = add_format({border: 0, align: 'right', size: 16, bold: 700})
    @logo_format.set_text_wrap
  end

  def average_format
    add_format({align: 'center', bg_color: 'white'})
  end

  def header_format
    {
        odd:  {align: 'center', bg_color: @bg_blue_color, color: 'black'},
        even: {align: 'center', bg_color: 'white',        color: 'black'}
    }
  end

  def prices_format
    {
        odd:  {align: 'center', bg_color: @bg_blue_color, color: @font_gray_color, font: 'Arial', num_format: '# ##0.00'},
        even: {align: 'center', bg_color: 'white',        color: @font_gray_color, font: 'Arial', num_format: '# ##0.00'}
    }
  end

  def set_stock_formats
    @color_index   = 55
    @stock_formats = {}
    @stock_with_currency_formats = {}
    Stock::ALL.each do |status_id|
      set_color('bgcolor', '#' + Stock::COLORS[status_id])
      @stock_formats[status_id] = add_format({
                                                 bg_color:    @bgcolor,
                                                 color:       (Stock::BOOKED_STATUSES.include?(status_id) ? :silver : :black),
                                                 border:      1,
                                                 num_format: '# ##0',
                                                 align:      'center'
                                             })

      @stock_with_currency_formats[status_id] = {}
      COUNTRIES.each do |country_id, country_hash|
        format = @workbook.add_format
        format.copy(@stock_formats[status_id])
        format.set_num_format("# ##0#{country_hash[:money]}")
        @stock_with_currency_formats[status_id][country_id] = format
      end
      @color_index += 1
    end
    @color_index
  end

  def set_color(variable, color)
    instance_variable_set("@#{variable.to_s}", @workbook.set_custom_color(@color_index, color))
  end

  #setting @header_gray
  #setting @bg_blue_color
  #setting @font_gray_color
  def set_custom_colors
    set_color('header_gray',     '#555555')
    @color_index += 1
    set_color('bg_blue_color',   '#ccffcc')
    @color_index += 1
    set_color('font_gray_color', '#3c3c3c')
  end

  def set_cell_formats_props(header_format, prices_format)
    @cell_formats_props = {
        price_sell:     prices_format,
        price_discount: prices_format,

        address: formats_props('left',   @font_gray_color, 14),
        side:    formats_props('center', @font_gray_color, 14),
        link:    formats_props('left',   'blue',          nil)
    }

    keys = [:num, :construction_id, :surface_id, :owner_code, :owner, :city, :supertype, :type, :grp, :ots, :stock_updated_at]
    keys.each { |k| @cell_formats_props[k] = header_format }
  end

  def formats_props(align, color, size)
    {
        odd:  {align: align, bg_color: @bg_blue_color,  color: color, size: size},
        even: {align: align, bg_color: 'white',         color: color, size: size}
    }
  end

  def set_cell_formats
    @cell_formats = {}
    @cell_formats_props.each do |id, props|
      @cell_formats[id] = {
          odd:  add_format(props[:odd]),
          even: add_format(props[:even])
      }
    end

    @cell_formats[:link][:even].set_text_wrap
    @cell_formats[:link][:odd].set_text_wrap
  end

  def set_average_formats(average_format)
    @average_formats = {}
    @average_formats[:label] = add_format({align: 'left', bg_color: 'white', bold: 700})
    @average_formats[:side], @average_formats[:grp], @average_formats[:ots] = average_format, average_format, average_format
  end
end