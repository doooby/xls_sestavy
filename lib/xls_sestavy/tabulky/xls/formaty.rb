# encoding: utf-8
module XLSSestavy
  module Xls
    module Formaty

      #do uchovávaného pole zapíše pod daným symbolem nový formát využívaný ostatními metodami
      def add_format(symbol, *args)
        @formaty = {} unless defined? @formaty
        @formaty[symbol] = @wb.add_format *args
      end

      def add_altered_format(symbol, format, zmeny_hash)
        @formaty = {} unless defined? @formaty
        @formaty[symbol] = alter_format format, zmeny_hash
      end

      # vytáhne vytvořený formát podle symbolu
      def get_format(symbol=nil)
        @formaty = {} unless defined? @formaty
        unless symbol
          @format = :default unless defined? @format
          symbol = @format
        end
        f = @formaty[symbol]
        return f if f
        add_default_format symbol
      end

      def alter_format(format, zmeny_hash)
        f = @wb.add_format
        f.copy format
        f.set_format_properties zmeny_hash
        f
      end

      #aktivní format je využíván dalšímí metodami
      def set_aktivni_format(format)
        @format = format
      end

      #definice defaultních formátů (voláno z add_format)
      def add_default_format(symbol)
        case symbol
          when :sestava_nadpis
            add_format symbol, bold: 1, size: 15, bg_color: 52, align: 'left'
          when :sestava_nadpis2
            add_format symbol, size: 12, bold: 1, align: 'left'
          when :sestava_info
            add_format symbol, italic: 1
          when :radek_zahlavi
            add_format symbol, bold: 1, border: 1, text_wrap: 1, align: 'center', bg_color: 22
          when :radky_dat
            add_format symbol, text_wrap: 1, border: 1
          when :radky_dat_stred
            add_format symbol, text_wrap: 1, border: 1, align: 'center'
          when :radek_souctu
            add_format symbol, align: 'right', bg_color: 42, bold: 1, border: 1
          when :default
            add_format symbol
          else
            raise "Formát :#{symbol} není definován"
        end
      end

    end
  end
end