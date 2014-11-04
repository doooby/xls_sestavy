# encoding: utf-8
require "xls_sestavy/tabulky/tabulka"
require "xls_sestavy/tabulky/sloupce/sloupec"
require "xls_sestavy/tabulky/sloupce/preddefinovany_sloupce"
require "xls_sestavy/tabulky/sloupce/definice_sloupcu"
require "xls_sestavy/tabulky/promenne_radku"

module XLSSestavy
  class Sestava

    def self.nazev
      @nazev ||= 'bezejména'
    end

    def self.nastav_nazev(txt)
      @nazev = txt.freeze
    end

    attr_reader :params

    def initialize(params=nil)
      @params = zpracuj_params(params).freeze
      @cas_vytvoreni = Time.zone.now.freeze
    end

    def zpracuj_params(params)
      params || {}
    end

    def zpracuj
      vypracovani
      block_given? ? yield : po_vypracovani
    end

    # přepsáno v podtřídách
    def vypracovani; end

    # přepsáno v podtřídách
    def po_vypracovani; end

    def to_s
      self.class.nazev.gsub ' ','_'
    end

  end
end