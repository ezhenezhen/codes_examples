#-*- encoding : utf-8 -*-
class FormsController < ApplicationController
  before_filter :detect_model

  MODELS = {
      managers:  { redirect_to_model: MessageToManager,
                   persisted: true,
                   subject: 'Быстрое сообщение от клиента с сайта',
                   success_message: 'Мы свяжемся с вами в ближайшее время.'
      },

      agents:    { redirect_to_model: Agent,
                   persisted: false,
                   subject: 'New agent registration',
                   success_message: 'Ваше сообщение отправлено'
      },

      companies: { redirect_to_model: Company,
                   persisted: false,
                   subject: 'New company registration',
                   success_message: 'Ваше сообщение отправлено'
      }
  }

  def new
    @object = @model[:redirect_to_model].new
    get_managers_ids
    render template_name
  end

  def create
    @object = @model[:redirect_to_model].new(params[:form])
    get_managers_ids
    valid = ((@model[:persisted] && @object.save) || (!@model[:persisted] && @object.valid?))
    valid ? success : prompt_to_fill_in_all_fields
  end

  def done
    render 'forms/done'
  end

  private

  def detect_model
    @model_name = params[:model]
    @model = MODELS[@model_name.to_sym]
    raise_404 unless @model
  end

  def send_email
    if @model[:redirect_to_model] == MessageToManager
      manager = Manager.find_by_internal_number(params[:form][:manager_id])
      ManagerMailer.message_to_manager(params, manager).deliver
    else
      FormMailer.email_form(params, @model[:subject]).deliver
    end
  end

  def success
    flash[:notice] = "Спасибо! #{@model[:success_message]}"
    send_email
    redirect_to done_path
  end

  def prompt_to_fill_in_all_fields
    flash.now[:error] = 'Пожалуйста, заполните все обязательные поля'
    render template_name
  end

  def template_name
    @model[:redirect_to_model].to_s.underscore
  end

  def get_managers_ids
    return unless @model[:redirect_to_model] == MessageToManager
    @manager_numbers = Manager.with_internal_number.order(:internal_number).pluck(:internal_number)
  end
end