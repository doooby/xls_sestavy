# encoding: utf-8
module XLSSestavy
  class Sloupec

    attr_accessor :tabulka

    # argumenty, fce, xls_format, datovy_typ, radek_souctu, sirka_sloupce
    def initialize(zahlavi, opts={}, &block)
      @options = opts.merge! zahlavi: zahlavi, fce: (opts[:fce]||block), argumenty: (opts[:argumenty]||[])
    end

    def zahlavi
      raise 'Není přiřazena tabulka.' unless @tabulka
      @zahlavi ||= I18n.translate(@options[:zahlavi], {default: @options[:zahlavi]}.merge!(@tabulka.sestava.params))
    end

    def argumenty_array(objekt)
      raise 'Není přiřazena tabulka.' unless @tabulka
      params = @tabulka.sestava.params
      if @tabulka.promenne
        promenne = @tabulka.promenne.pro_objekt(objekt)
        @options[:argumenty].map{|arg_sym| params[arg_sym] || promenne[arg_sym]}
      else
        @argumenty_array ||= @options[:argumenty].map{|arg_sym| params[arg_sym]}
      end
    end

    def hodnota_pro(objekt)
      @fce ||= @options[:fce]
      if @fce.class==Proc
        @fce.call(objekt, *argumenty_array(objekt))
      elsif objekt.respond_to? @fce
        objekt.send(@fce, *argumenty_array(objekt))
      else
        raise "Není definováno pro @fce=#{@fce.class}:#{@fce}"
      end
    end

    def opt(klic)
      @options[klic]
    end

  end
end