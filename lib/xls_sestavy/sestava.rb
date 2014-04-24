# encoding: utf-8
require 'tempfile'
require 'xls_sestavy/excel_formaty'
require 'xls_sestavy/excel_tabulky'


module XLSSestavy
  class Sestava
    NAZEV = 'Prázdná sestava'

    include ExcelTabulky
    include ExcelFormaty

    attr_reader :argumenty

    def initialize(args={})
      @argumenty = ArgumentySestavy.new args
    end

    def vytvor_soubor
      f = Tempfile.new self.class::NAZEV.gsub(' ','_'), Rails.root.join('tmp')
      @wb = WriteExcel.new f
      vypracuj_sestavu
      @wb.close
      block_given? ? yield(f) : po_vypracovani(f)
    end

    def vypracuj_sestavu; end

    def po_vypracovani(f); end

    def nazev_souboru
      to_s + '.xls'
    end

    def to_s
      "#{self.class::NAZEV.gsub(' ','_')}_#{@argumenty.to_s}"
    end

  end
end