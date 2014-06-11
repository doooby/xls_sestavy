# encoding: utf-8
module XLSSestavy
  class Tabulka

    def initialize(sestava, args={}, &block)
      @sestava = sestava
      @sloupce = args[:sloupce] || []
      @args = args
      yield self if block_given?
    end

    def pridej(*args, &block)
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