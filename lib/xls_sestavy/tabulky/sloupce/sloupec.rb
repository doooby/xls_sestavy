# encoding: utf-8
module XLSSestavy
  class Sloupec

    attr_accessor :tabulka

    def initialize(zahlavi, opts={}, &block)
      @options = opts.merge! zahlavi: zahlavi, fce: (opts[:fce] || block)
    end

    def zahlavi
      raise 'Není přiřazena tabulka.' unless @tabulka
      @zahlavi ||= I18n.translate(@options[:zahlavi], {default: @options[:zahlavi]}.merge!(@tabulka.sestava.params))
    end

    def argumenty_array
      return @argumenty_array if defined? @argumenty_array
      raise 'Není přiřazena tabulka.' unless @tabulka
      params = @tabulka.sestava.params
      @argumenty_array = (@options[:argumenty] || []).map{|arg_sym| params[arg_sym]}
    end

    def hodnota_pro(objekt)
      @fce ||= @options[:fce]
      if @fce.class==Proc
        @fce.call(objekt, *argumenty_array)
      elsif objekt.respond_to? @fce
        objekt.send(@fce, *argumenty_array)
      elsif nil
        nil
      else
        raise "Není definováno pro @fce=#{@fce.class}:#{@fce}"
      end
    end

    def opt(klic)
      @options[klic]
    end

  end
end