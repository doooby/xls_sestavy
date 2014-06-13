# encoding: utf-8
module XLSSestavy
  class Tabulka

    attr_reader :sestava

    def initialize(sestava, args={}, &block)
      @sestava = sestava
      @args = args

      @sloupce = args[:sloupce] || []
      yield self if block_given?
      @sloupce.each{|s| s.tabulka = self}
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

  end
end