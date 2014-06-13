# encoding: utf-8
require 'tempfile'
require 'writeexcel'
require "xls_sestavy/xls/xls"
require "xls_sestavy/xls/formaty"
require "xls_sestavy/xls/zapisovani"
require "xls_sestavy/tabulky/tabulka_xls"

module XLSSestavy
  class SestavaXls < Sestava

    include XLSSestavy::Xls::Formaty
    include XLSSestavy::Xls::Zapisovani

    def zpracuj
      f = Tempfile.new to_s, Rails.root.join('tmp')
      @wb = ::WriteExcel.new f
      vypracovani
      @wb.close
      block_given? ? yield(f) : po_vypracovani(f)
    end

    # přepsáno v podtřídách
    def vypracovani; end

    # přepsáno v podtřídách
    def po_vypracovani(f); end

    def to_s
      self.class.nazev.gsub ' ','_'
    end

    def nazev_souboru
      to_s + '.xls'
    end

    # vytvoří list, zapíše jej do @ws a vrátí jej. mezi tím případně provede předaný blok
    def vytvor_list(nazev)
      @ws = @wb.add_worksheet nazev
      yield @ws if block_given?
      @ws
    end

  end
end