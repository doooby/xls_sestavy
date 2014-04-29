# encoding: utf-8
module XLSSestavy
  class RadaSloupcu

    attr_reader :sloupce

    def initialize(sloupce=nil, argumenty=nil)
      @sloupce = sloupce || []
      yield self if block_given?
      if @sloupce.class!=Array && @sloupce.any?{|s| !s.kind_of? Sloupec}
        raise 'RadaSloupcu musí být inicializovány polem Sloupců'
      end
      @sloupce.each{|s| s.nastav_argumenty argumenty} if argumenty
      @seznam = {}
      aktualizuj_seznam
    end

    def pridej(*args, &block)
      if args.first.is_a? Sloupec
        @sloupce << args.first
      else
        @sloupce << Sloupec.new(*args, &block)
      end
    end

    def [](klic)
      index = @seznam[klic.to_sym]
      index ? @sloupce[index] :nil
    end

    def aktualizuj_seznam
      @seznam = {}
      @sloupce.each_with_index do |s,i|
        @seznam[s.klic] = i if s.class==PreddefinovanySloupec
      end
    end

  end
end