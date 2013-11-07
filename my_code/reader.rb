# -*- encoding : utf-8 -*-
class FeedReaders::ArtMaster < SmartFeedReader
  self.domain   = 'non-existent-domain'
  self.owner_id = 1520
  #self.debug    = true

  self.recognize = {
      by:  :constructions,
      key: :code_from_owner
  }

  self.notify_after = 5.days
  self.notify_name  = 'Дмитрий'
  # self.notify_email = ''

  def run
    for_first_sheet do

      columns   address:             '\A\s*АДРЕС\s*\z',
                code_unprocessed:    '\A\s*№\s*\z',
                price_unprocessed:   '\A\s*ПРАЙС\s*\z',
                link_to_source_page: '\A\s*ФОТО\s*\z',
                side:                '\A\s*СТОРОНА\s*\z'

      # колонки цены и стороны определяются под одним и тем же номером из-за
      # их расположения и того, что они объединены. исправляем это:
      @columns[:price_unprocessed] += 1 if @columns[:side] == @columns[:price_unprocessed]

      unmerge_cells_in_columns :address, :code_unprocessed, :side, :price_unprocessed, :link_to_source_page

      readers   code:  -> { get(:code_unprocessed).to_s.squish.sub(/\.0$/, '') },
                price: -> { get_price_warn_if_zero }

      intervals type:                :regexp,
                match:               custom_interval_matchers

      prepare   self.class.recognize[:by],
                key:                 self.class.recognize[:key]

      read      correspond:          :construction,
                key_column:          :code,
                find_surface:        :by_side_name

      mark_sold

    end
  end

  private

  def get_price_warn_if_zero
    price = get(:price_unprocessed)
    send_warning(self, 'Цена определилась как нулевая. Возможна ошибка при определении колонки цены') if price.to_i == 0
    price
  end

  def detect_status(row, column, cell, color, construction, surface)
    return Stock::SOLD     if %w(red xls_color_21).include? color
    return Stock::RESERVED if %w(yellow).include? color
    Stock::FREE            if %w(white border).include? color
  end

  def custom_interval_matchers
    %w(
        \A\s*01\s*\z
        \A\s*02\s*\z
        \A\s*03\s*\z
        \A\s*04\s*\z
        \A\s*05\s*\z
        \A\s*06\s*\z
        \A\s*07\s*\z
        \A\s*08\s*\z
        \A\s*09\s*\z
        \A\s*10\s*\z
        \A\s*11\s*\z
        \A\s*12\s*\z
      )
  end
end