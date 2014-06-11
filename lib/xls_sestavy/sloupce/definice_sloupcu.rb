# encoding: utf-8
module XLSSestavy
  class DefiniceSloupcu < RadaSloupcu

    def self.[](klic)
      @instance ||= new
      @instance[klic]
    end

    def initialize
      @seznam = {}
      definice
    end

    def definice; end # tato metoda bude přepsána

    def definuj(*args, &block)
      s = PreddefinovanySloupec.new(*args, &block)
      @seznam[s.klic] = s
    end

    def [](klic)
      @seznam[klic]
    end

  end
end