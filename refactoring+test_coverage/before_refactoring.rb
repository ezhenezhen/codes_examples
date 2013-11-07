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
      {                                          :description => 'Добавить колонки:'}, # group header
      {:id => :enable_our_constructions_numbers, :description => 'Наш № РК', :hide_from_owners => true},
      {:id => :enable_our_surfaces_numbers,      :description => 'Наш № поверхности', :default => true, :hide_from_owners => true},
      {:id => :enable_owner_codes,               :description => '№ поверхности владельца', :hide_from_clint => true},
      {:id => :enable_supertype,                 :description => 'Тип'},
      {:id => :enable_types,                     :description => 'Формат'},
      {:id => :enable_grp,                       :description => 'GRP'},
      {:id => :enable_ots,                       :description => 'OTS'},
      {:id => :calculate_average,                :description => 'Средние показатели'},
      {:id => :with_owners,                      :description => 'Владелец', :hide_from_owners => true, :hide_from_clint => true},
      {:id => :city,                             :description => 'Город'},
      {:id => :add_stock_date,                   :description => 'Дата обновления статуса'},
      {:id => :hide_all_prices_and_states,       :description => 'Скрыть все статусы и цены'},
      {:id => :skip_labels,                      :description => 'Пропустить комментарии'},
      {                                          :description => 'Настройки:', :hide_from_owners => true},
      {:id => :hide_prices,                      :description => 'Режим лояльности (скрывать цены)', :hide_from_owners => true, :hide_from_clint => true},
      {:id => :create_surfaces_list,             :description => 'Создать страницу АП', :hide_from_owners => true},
      {:id => :start_from_current_interval,      :description => 'Начать с текущего месяца/декады', :default => lambda{Time.now.day <= 5}, :hide_from_owners => true},
      {:id => :show_prices_for_sold,             :description => 'Показывать цены для проданных'},
      {                                          :description => 'Ставить ссылку на:', :hide_from_owners => true},
      {:id => :link_to_surfaces_list,            :description => 'Страницу АП', :hide_from_owners => true},
      {:id => :link_to_surfaces_pages,           :description => 'Cтраницу стороны', :hide_from_owners => true},
  ]

  EMAIL = 'xxx'
  PHONE = 'xxx'

  HEADER_ROWS = 3
  def initialize(opts = {})
    @options = opts
    @surfaces = []
    define_columns

    # Deserialize discounts
    @discounts = {}
    if @options[:discounts]
      @options[:discounts].split(',').each{|d| dd = d.split(':'); @discounts[dd.first.to_i] = dd.last.to_i}
    end

    opts[:protocol_host_port] ||= "http://#{PRODUCTION_HOST}"
  end

  def self.options_for(vint, clint)
    opts = []
    OPTIONS.each do |o|
      next if o[:hide_from_owners] && vint
      next if o[:hide_from_owners] && clint
      opts << o
    end
    opts
  end

  def define_columns
    @columns = {}
    @columns_widths = []
    @logo_join_columns = 1
    @shift_status_legend = 0

    # Defining for decades
    @columns[:d] ||= []
    @columns[:d] << {:label => '№'}
    @columns_widths << 6

    if(@options[:enable_our_constructions_numbers])
      @columns[:d] << {:label => '№ РК'}
      @columns_widths << 16
      @logo_join_columns += 1
    end

    if(@options[:enable_our_surfaces_numbers])
      @columns[:d] << {:label => '№ Пов-ти'}
      @columns_widths << 16
      @logo_join_columns += 1
    end

    if(@options[:enable_owner_codes])
      @columns[:d] << {:label => '№ Влад-ца'}
      @columns_widths << 16
      @logo_join_columns += 1
    end

    if(@options[:with_owners])
      @columns[:d] << {:label => 'Влад-ц'}
      @columns_widths << 16
      @logo_join_columns += 1
      @shift_status_legend += 1
    end

    if(@options[:city])
      @columns[:d] << {:label => 'Город'}
      @columns_widths << 16
      @logo_join_columns += 1
      @shift_status_legend += 1
    end

    if(@options[:enable_supertype])
      @columns[:d] << {:label => 'Тип'}
      @columns_widths << 16
      @logo_join_columns += 1
      @shift_status_legend += 1
    end

    if(@options[:enable_types])
      @columns[:d] << {:label => 'Формат'}
      @columns_widths << 16
      @logo_join_columns += 1
      @shift_status_legend += 1
    end

    @columns[:d] << {:label => 'Адрес'}
    @columns_widths << 79

    @columns[:d] << {:label => 'Сторона'}
    @columns_widths << 15

    if(@options[:enable_grp])
      @columns[:d] << {:label => 'GRP'}
      @columns_widths << 10
      @logo_join_columns += 1
      @shift_status_legend += 1
    end

    if(@options[:enable_ots])
      @columns[:d] << {:label => 'OTS'}
      @columns_widths << 10
      @logo_join_columns += 1
      @shift_status_legend += 1
    end

    @columns[:d] << {:label => 'Ссылка на фото, карту'}
    @columns_widths << 25

    if(@options[:add_stock_date])
      @columns[:d] << {:label => 'Дата обновления'}
      @columns_widths << 16
      @logo_join_columns += 1
      @shift_status_legend += 1
    end

    # Defining for months
    @columns[:m] ||= []
    @columns[:m] << {:label => '№'}
    @columns[:m] << {:label => '№ РК'}      if @options[:enable_our_constructions_numbers]
    @columns[:m] << {:label => '№ Пов-ти'}  if @options[:enable_our_surfaces_numbers]
    @columns[:m] << {:label => '№ Влад-ца'} if @options[:enable_owner_codes]
    @columns[:m] << {:label => 'Влад-ц'}    if @options[:with_owners]
    @columns[:m] << {:label => 'Город'}     if @options[:city]
    @columns[:m] << {:label => 'Тип'}       if @options[:enable_supertype]
    @columns[:m] << {:label => 'Формат'}    if @options[:enable_types]
    @columns[:m] << {:label => 'Адрес'}
    @columns[:m] << {:label => 'Сторона'}
    @columns[:m] << {:label => 'GRP'}       if @options[:enable_grp]
    @columns[:m] << {:label => 'OTS'}       if @options[:enable_ots]
    @columns[:m] << {:label => 'Ссылка на фото, карту'}
    @columns[:m] << {:label => 'Дата обновления'} if @options[:add_stock_date]
  end

  # Add surface or list of surfaces to export
  def add(surface)
    case surface
      when Surface then @surfaces << surface
      when Array then surface.each{|s| add(s)}
      else raise "Can add only surfaces or array of surfaces"
    end
  end

  # Log this export to DB
  def log_export
    sql = <<-SQL
      INSERT INTO surfaces_excel_exports(manager_id, options, surfaces_ids, created_at) VALUES (#{@options[:manager].is_a?(Manager) ? @options[:manager].id : 'NULL'}, '#{@options.to_json}', '#{@surfaces.map(&:id) * ','}', now())
    SQL
    ActiveRecord::Base.connection.execute(sql)
  #rescue
    # nothing
  end

  def do!
    raise "No surfaces to export" if @surfaces.blank?
    log_export
    surfaces_by_type = Surface.separate_by_tariff(@surfaces)
    filename = Tempfile.new('surfaces_excel_export').path + '.xls'

    @workbook = WriteExcel.new(filename)
    worksheet  = @workbook.add_worksheet(Date.today.strftime("%d.%m.%Y"))
    worksheet.set_zoom(75)
    prepare_formats

    estimate_rows = @surfaces.length
    estimate_rows += (surfaces_by_type.length - 1) * 2 # two rows between types
    estimate_rows += surfaces_by_type.length           # row between table and labels
    estimate_rows += surfaces_by_type.length * 2       # 2 rows for each headers
    surfaces_by_type.keys.each do |p|
       t = Tariff.class_by_prefix(p)
       estimate_rows += [t::EXPORT_LABELS.length + t::EXPORT_LINKS.length, 4].max + 1

       # onw row spacer and one row for average line
       estimate_rows += 2 if @options[:calculate_average]
    end
    (0..(HEADER_ROWS - 1)).to_a.each{|i| worksheet.set_row(i, 25)}
    (HEADER_ROWS..(HEADER_ROWS + estimate_rows)).to_a.each{|i| worksheet.set_row(i, 21)}

    # Insert logo
    worksheet.merge_range(0, 0, 1, @logo_join_columns, "Телефон: #{@options[:phone] || PHONE}\r\nEmail: #{@options[:email] || EMAIL}", @logo_format)

    logo = (@options[:logo] || File.join(Rails.root, %w(public i logo.png)))
    worksheet.insert_image('A1', logo, 10, 10, 1, 1.1)
    worksheet.write(0, @logo_join_columns + 2, "По состоянию на #{Date.today.strftime('%d.%m.%Y')}", @timestamp)

    # Setting columns widths
    total_columns = @columns_widths.length
    @columns_widths.each_with_index do |width, index|
      worksheet.set_column(index, index, width)
    end

    max_intervals_shown = surfaces_by_type.keys.map{|k| Tariff.class_by_prefix(k).show_intervals}.max
    worksheet.set_column(total_columns, max_intervals_shown + total_columns - 1, 20)

    @surfaces_list = SurfacesList.create(:manager => @options[:manager]) if @options[:create_surfaces_list] || @options[:link_to_surfaces_list]

    row = HEADER_ROWS
    surfaces_by_type.each do |tariff_prefix, surfaces|

      tariff = Tariff.class_by_prefix(tariff_prefix)
      intervals = []
      # Start from current or next interval
      shift_number = (@options[:start_from_current_interval] ? 0 : 1)
      tariff.show_intervals.times{ |i| intervals << tariff.now.next(i + shift_number)}

      # Writing headers and joining cells in headers
      @columns[tariff_prefix].each_with_index do |column, index|
        worksheet.merge_range(row, index, row + 1, index, column[:label], @header_cell_format)
      end
      worksheet.merge_range(row, total_columns, row, total_columns + intervals.length - 1, 'Периоды', @header_cell_format)
      row += 1

      # Writing interval headers
      intervals.each_with_index do |interval, i|
        worksheet.write(row, total_columns + i, interval.text_name, @header_cell_format_intervals)

        # Set width for interval columns
        worksheet.set_column(total_columns + i, total_columns + i, 15)
      end

      # Define columns order
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
      dataset = []
      prices_dataset = {}

      surfaces.each_with_index do |surface, index|
        @surfaces_list.add(surface) if @options[:create_surfaces_list] || @options[:link_to_surfaces_list]
        row += 1

        num = index + 1
        data = {
            :num              => num,
            :construction_id  => surface.construction.id,
            :surface_id       => surface.id,
            :owner_code       => surface.owner_code,
            :owner            => surface.construction.owner.name,
            :city             => surface.construction.region.name,
            :supertype        => layout_supertype_singular(surface.construction.supertype),
            :type             => surface.construction.type_name,
            :grp              => surface.grp,
            :ots              => surface.ots,
            :stock_updated_at => (dt = surface.stock_updated_at; dt ? dt.strftime("%d.%m.%Y") : ''),
            :address          => layout_address(surface),
            :side             => layout_side(surface),
            :link             => link(surface)
        }
        dataset << data

        order.each_with_index do |id, idx|
          #f = @cell_formats[id][num.odd? ? :odd : :even]
          f = @cell_formats[id][:even]
          if(id == :link)
            worksheet.write_url(row, idx, data[id], data[id], f)
          else
            worksheet.write(row, idx, data[id], f)
          end
        end

        unless @options[:hide_all_prices_and_states]
          col = total_columns
          prices = surface.all_prices_hash
          intervals.each do |interval|
            prices_dataset[interval] ||= []
            state = surface.stock_at(interval)

            value = ''
            if state == Stock::SOLD && !(@options[:show_prices_for_sold])
              value = 'Продано'
            else
              value = (@options[:hide_prices].blank? ? prepare_price_value(prices, interval, surface.construction.owner_id) : '')
              prices_dataset[interval] << value if value.present?
            end

            worksheet.write(row, col, value, @stock_formats[state])
            col += 1
          end
        end
      end

      # Writing average values if required
      if @options[:calculate_average]
        row += 2
        # Label and average mediaindicators
        worksheet.write(row, order.index(:address), 'Средние показатели', @average_formats[:label])
        worksheet.write(row, order.index(:grp), '%0.2f' % dataset.map{|d| d[:grp]}.compact.avg, @average_formats[:grp]) if @options[:enable_grp]
        worksheet.write(row, order.index(:ots), '%0.2f' % dataset.map{|d| d[:ots]}.compact.avg, @average_formats[:ots]) if @options[:enable_ots]

        #Calculating and writing oercentage of A sides
        sides = dataset.map{|d| d[:side]}.compact
        a_sides = sides.select{|s| Surface.is_side_a?(s)}
        a_sides_percentage = ((a_sides.length.to_f / sides.length.to_f) * 100).to_i
        worksheet.write(row, order.index(:side), "Сторон А: #{a_sides_percentage}%", @average_formats[:side])

        # Writing average prices for each month
        col = total_columns
        intervals.each do |interval|
          next unless prices_dataset[interval]
          avg = prices_dataset[interval].avg
          worksheet.write(row, col, avg > 0 ? avg.to_i : '', @stock_formats[Stock::UNKNOWN])
          col += 1
        end
      end

      # Writing labels and links under the table
      unless @options[:skip_labels]
        row += 2
        first_label_row = row
        (tariff::EXPORT_LABELS + tariff::EXPORT_LINKS).each do |label|
          if(label.is_a?(Array))
            worksheet.write_url(row, 1, label.first, label.second, @link_format)
          else
            worksheet.write(row, 1, label, @labels_format)
          end
          row += 1
        end
        last_label_row = row

        # Writing statuses legend
        row = first_label_row # + 2
        [Stock::RESERVED, Stock::FREE, Stock::SOLD, Stock::UNKNOWN].each do |status|
          worksheet.write(row, @shift_status_legend + 5, '', @stock_formats[status])
          worksheet.write(row, @shift_status_legend + 6, ' - ' + STOCK_LEGENDS[status], @stock_legend_format)
          row += 1
        end
      end

      # Adding link to Active Sales page
      if @options[:create_surfaces_list]
        worksheet.write_url(1, @logo_join_columns + 2, @surfaces_list.link(@options[:protocol_host_port]), 'Ваша адресная программа на карте', @link_format)
      end

      row += 3
    end

    @workbook.close
    @filename = filename
  end

  def surfaces_list_link
    @surfaces_list.link
  end

  def self.clear_parameters(params, manager)
    h = {}
    OPTIONS.each do |opt|
      next if params["excel_option_#{opt[:id]}"].blank?
      next if opt[:hide_from_owners] && manager.department.owner  # Don't pass hidden from owners parameters if current manager is from owner
      next if opt[:hide_from_clint]  && manager.clint?            # Don't pass hidden from clint managers parameters if current manager is clint
      h[opt[:id]] = true
    end
    h
  end

  private

  def prepare_price_value(prices, interval, owner_id)
    p = prices[interval.id]
    return '' if p.blank?
    price = p[:sell] || p[:base]

    # Enable discount if present for this owner
    if price && @discounts[owner_id]
      delta = (@options[:discount_unit] == '%' ? price * @discounts[owner_id] / 100 : @discounts[owner_id])
      price = (@options[:discount_direction] == '-' ? price - delta : price + delta)
    end

    price
    #(p[:discount].present? && p[:discount] > 0 && p[:discount] < p[:sell]) ? p[:discount] : p[:sell]
  end

  def layout_address(surface)
    surface.construction.address
  end

  def layout_side(surface)
    surface.name.gsub('Сторона ','')
  end

  def price(val)
    val == 0 ? ' ' : number_to_currency(val, :unit => "", :precision => 0, :delimiter => " ", :format => "%n")
  end

  # Prepare link for given surface depend on parameters
  def link(s)
    protocol_host_port = @options[:protocol_host_port]

    if @options[:link_to_external_site]
      s.exteral_site_link(protocol_host_port)
    elsif @options[:hide_prices]
      # Links to surface page without prices. For loyal clients
      s.link_to_noprice_page(protocol_host_port)
    elsif @options[:link_to_surfaces_list]
      # Links to SurfacesList page with anchor to exact surface
      @surfaces_list.link_for_surface(s, protocol_host_port)
    elsif @options[:link_to_surfaces_pages]
      # Links to surface page with all details (with price)
      s.link(protocol_host_port)
    else
      # By default insert links to index map with construction details in side panel
      s.construction.link_to_page(protocol_host_port)
    end
  end

  def prepare_formats
    @default_cell_props = {:border => 1, :valign => 'vcenter', :font => 'Arial', :size => 11}

    color_index = 55
    @stock_formats = {}
    Stock::ALL.each do |status_id|
      bgcolor = @workbook.set_custom_color(color_index, '#' + Stock::COLORS[status_id])
      @stock_formats[status_id] = @workbook.add_format(@default_cell_props.merge(
          :bg_color => bgcolor,
          :color => (Stock::BOOKED_STATUSES.include?(status_id) ? :silver : :black),
          :border => 1,
          :num_format => '# ##0p',
          :align => 'center'
      ))
      color_index += 1
    end

    header_gray = @workbook.set_custom_color(color_index, '#555555')
    header_cell_format_props = {
        :border  => 1,
        :valign  => 'vcenter',
        :align   => 'center',
        :bg_color => header_gray,
        :color => 'white',
        :bold => 700,
        :size => 11,
        :text_wrap => 1
    }

    @header_cell_format = @workbook.add_format(header_cell_format_props)
    # Create totally copy of usual styles. Don't know why but Excel from Office XP crash on exports generated with the same styles for intervals headers cells
    # So that is just workaround for Excel XP
    @header_cell_format_intervals = @workbook.add_format(header_cell_format_props)

    color_index += 1

    @logo_format = @workbook.add_format(
                             :border  => 0,
                             :valign  => 'vcenter',
                             :align   => 'right',
                             :font => 'Arial',
                             :size => 16,
                             :bold => 700
                           )
    @logo_format.set_text_wrap

    @timestamp = @workbook.add_format(
                             :border  => 0,
                             :valign  => 'vcenter',
                             :align   => 'left',
                             :font => 'Arial',
                             :size => 13
                           )

    @labels_format = @workbook.add_format(
                             :border  => 0,
                             :valign  => 'vcenter',
                             :align   => 'left',
                             :size => 13,
                             :font => 'Arial'
                           )

    @labels_link_format = @workbook.add_format(
                             :border  => 1,
                             :border_color => 'gray',
                             :valign  => 'vcenter',
                             :align   => 'left',
                             :size => 13,
                             :font => 'Arial',
                             :color => 'blue'
                           )

    @link_format = @workbook.add_format(
                             :border  => 0,
                             :valign  => 'vcenter',
                             :align   => 'left',
                             :size => 13,
                             :underline => 1,
                             :font => 'Arial',
                             :color => 'blue'
                           )

    @stock_legend_format = @workbook.add_format(
                             :border  => 1,
                             :border_color => 'black',
                             :valign  => 'vcenter',
                             :align   => 'left',
                             :bg_color => 'white',
                             :size => 11,
                             :font => 'Arial'
                           )

    bg_blue_color = @workbook.set_custom_color(color_index, '#ccffcc')
    color_index += 1
    font_gray_color = @workbook.set_custom_color(color_index, '#3c3c3c')
    color_index += 1

    @cell_formats_props = {
        :num => {
            :odd => {:align => 'center',:bg_color => bg_blue_color, :color => 'black'},
            :even => {:align => 'center',:bg_color => 'white', :color => 'black'}
        },
        :construction_id => {
            :odd => {:align => 'center',:bg_color => bg_blue_color, :color => 'black'},
            :even => {:align => 'center',:bg_color => 'white', :color => 'black'}
        },
        :surface_id => {
            :odd => {:align => 'center',:bg_color => bg_blue_color, :color => 'black'},
            :even => {:align => 'center',:bg_color => 'white', :color => 'black'}
        },
        :owner_code => {
            :odd => {:align => 'center',:bg_color => bg_blue_color, :color => 'black'},
            :even => {:align => 'center',:bg_color => 'white', :color => 'black'}
        },
        :owner => {
            :odd => {:align => 'center',:bg_color => bg_blue_color, :color => 'black'},
            :even => {:align => 'center',:bg_color => 'white', :color => 'black'}
        },
        :city => {
            :odd => {:align => 'center',:bg_color => bg_blue_color, :color => 'black'},
            :even => {:align => 'center',:bg_color => 'white', :color => 'black'}
        },
        :supertype => {
            :odd => {:align => 'center',:bg_color => bg_blue_color, :color => 'black'},
            :even => {:align => 'center',:bg_color => 'white', :color => 'black'}
        },
        :type => {
            :odd => {:align => 'center',:bg_color => bg_blue_color, :color => 'black'},
            :even => {:align => 'center',:bg_color => 'white', :color => 'black'}
        },
        :grp => {
            :odd => {:align => 'center',:bg_color => bg_blue_color, :color => 'black'},
            :even => {:align => 'center',:bg_color => 'white', :color => 'black'}
        },
        :ots => {
            :odd => {:align => 'center',:bg_color => bg_blue_color, :color => 'black'},
            :even => {:align => 'center',:bg_color => 'white', :color => 'black'}
        },
        :stock_updated_at => {
            :odd => {:align => 'center',:bg_color => bg_blue_color, :color => 'black'},
            :even => {:align => 'center',:bg_color => 'white', :color => 'black'}
        },
        :address => {
            :odd => {:align => 'left',:bg_color => bg_blue_color, :color => font_gray_color, :size => 14},
            :even => {:align => 'left',:bg_color => 'white', :color => font_gray_color, :size => 14}
        },
        :side => {
            :odd => {:align => 'center',:bg_color => bg_blue_color, :color => font_gray_color, :size => 14},
            :even => {:align => 'center',:bg_color => 'white', :color => font_gray_color, :size => 14}
        },
        :link => {
            :odd => {:align => 'left',:bg_color => bg_blue_color, :color => 'blue'},
            :even => {:align => 'left',:bg_color => 'white', :color => 'blue'}
        },
        :price_sell => {
            :odd => {:align => 'center',:bg_color => bg_blue_color, :color => font_gray_color, :font => 'Arial', :num_format => '# ##0.00'},
            :even => {:align => 'center',:bg_color => 'white', :color => font_gray_color, :font => 'Arial', :num_format => '# ##0.00'}
        },
        :price_discount => {
            :odd => {:align => 'center',:bg_color => bg_blue_color, :color => font_gray_color, :font => 'Arial', :num_format => '# ##0.00'},
            :even => {:align => 'center',:bg_color => 'white', :color => font_gray_color, :font => 'Arial', :num_format => '# ##0.00'}
        }
    }

    @cell_formats = {}
    @cell_formats_props.each do |id, props|
      @cell_formats[id] = {
          :odd => @workbook.add_format(@default_cell_props.merge(props[:odd])),
          :even => @workbook.add_format(@default_cell_props.merge(props[:even]))
      }
    end
    @cell_formats[:link][:even].set_text_wrap
    @cell_formats[:link][:odd].set_text_wrap

    @average_formats = {}
    @average_formats[:label] = @workbook.add_format(@default_cell_props.merge({:align   => 'left',   :bg_color => 'white', :bold => 700}))
    @average_formats[:side]  = @workbook.add_format(@default_cell_props.merge({:align   => 'center', :bg_color => 'white'}))
    @average_formats[:grp]   = @workbook.add_format(@default_cell_props.merge({:align   => 'center', :bg_color => 'white'}))
    @average_formats[:ots]   = @workbook.add_format(@default_cell_props.merge({:align   => 'center', :bg_color => 'white'}))
  end
end
