# encoding: utf-8
module XLSSestavy
  class Tabulka

    attr_reader :sestava, :promenne

    def initialize(sestava, args={}, &block)
      @sestava = sestava
      @args = args

      @sloupce = args[:sloupce] || []
      @promenne = args[:promenne]

      yield self if block_given?

      @sloupce.each{|s| s.tabulka = self}
      @promenne.pro_tabulku self if @promenne
    end

    def sloupec(*args, &block)
      s = if args.first.is_a? Sloupec
        args.first
      else
        Sloupec.new(*args, &block)
      end
      s.tabulka = self
      @sloupce << s
    end

    def promenna(klic, &block)
      @promenne ||= PromenneRadku.new
      @promenne.nastav klic, &block
    end

  end
end