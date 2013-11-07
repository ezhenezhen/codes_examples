# -*- encoding : utf-8 -*-
require 'spec_helper'

describe 'companies' do
  before :each do
    visit '/forms/companies'
  end

  it 'should not send empty form and show an error message' do
    click_button 'Отправить'
    page.should have_text('Пожалуйста, заполните все обязательные поля')
    page.find('#form_company_name_input').should have_text('не может быть пустым')
    page.find('#form_contact_input').should have_text('не может быть пустым')
    page.find('#form_phone_input').should have_text('не может быть пустым')
    page.find('#form_email_input').should have_text('не может быть пустым')
    page.find('#form_regions_input').should have_text('не может быть пустым')
    page.find('#form_quantity_input').should have_text('не может быть пустым')
    page.find('#form_available_input').should have_text('не может быть пустым')
    page.find('#form_pictures_input').should have_text('не может быть пустым')
    page.find('#form_maps_input').should have_text('не может быть пустым')
    page.find('#form_gids_input').should have_text('не может быть пустым')
  end

  it 'should check the validity of email and show an error message' do
    fill_in 'Название компании:', :with => 'Test Testing'
    fill_in 'Адрес:', :with => 'Address'
    fill_in 'Имя и фамилия контактного лица:', :with => 'Contact'
    fill_in 'Телефон:', :with => '8 800 555 9588'
    fill_in 'Email:', :with => 'test'
    fill_in 'Сайт:', :with => 'http:'
    fill_in 'В каких регионах (населённых пунктах, областях) расположены ваши конструкции?*', :with => 'region and city'
    fill_in 'Примерное количество конструкций:', :with => '333'
    choose 'form_available_yes'
    choose 'form_state_always_up_to_date'
    choose 'form_pictures_yes'
    choose 'form_maps_no'
    choose 'form_gids_no'
    choose 'form_column_yes'
    choose 'form_create_gids_no'
    fill_in 'Ваш комментарий', :with => 'some comment'
    click_button 'Отправить'
    page.find('#form_email_input').should have_text('имеет неверное значение')
  end

  it 'should send a form when all the required fields filled in' do
    fill_in 'Название компании:', :with => 'Test Testing'
    fill_in 'Адрес:', :with => 'Address'
    fill_in 'Имя и фамилия контактного лица:', :with => 'Contact'
    fill_in 'Телефон:', :with => '8 800 555 9588'
    fill_in 'Email:', :with => 'test@test.com'
    fill_in 'Сайт:', :with => 'http:'
    fill_in 'В каких регионах (населённых пунктах, областях) расположены ваши конструкции?*', :with => 'region and city'
    fill_in 'Примерное количество конструкций:', :with => '333'
    choose 'form_available_yes'
    choose 'form_state_always_up_to_date'
    choose 'form_pictures_yes'
    choose 'form_maps_no'
    choose 'form_gids_no'
    choose 'form_column_yes'
    choose 'form_create_gids_no'
    fill_in 'Ваш комментарий', :with => 'some comment'
    click_button 'Отправить'
    page.should have_content('Спасибо! Ваше сообщение отправлено')
  end

  describe 'agents' do
    before :each do
      visit '/forms/agents'
    end

    it 'should not send empty form and show an error message' do
      click_button 'Отправить'
      page.should have_text('Пожалуйста, заполните все обязательные поля')
      page.find('#form_name_input').should have_text('не может быть пустым')
      page.find('#form_age_input').should have_text('не может быть пустым')
      page.find('#form_phone_input').should have_text('не может быть пустым')
      page.find('#form_email_input').should have_text('не может быть пустым')
      page.find('#form_experience_input').should have_text('не может быть пустым')
    end

    it 'should check the validity of email and show an error message' do
      fill_in 'Имя и фамилия:', :with => 'Test Testing'
      fill_in 'Возраст:', :with => '22'
      fill_in 'Контактный телефон:', :with => '88005559588'
      fill_in 'Email:', :with => 'testasdflkj'
      fill_in 'Сколько лет вы работаете в рекламе:', :with => '2'
      click_button 'Отправить'
      page.find('#form_email_input').should have_text('имеет неверное значение')
    end

    it 'should show a confirmation when all the required fields filled in' do
      fill_in 'Имя и фамилия:', :with => 'Test Testing'
      fill_in 'Возраст:', :with => '22'
      fill_in 'Контактный телефон:', :with => '88005559588'
      fill_in 'Email:', :with => 'test@test.com'
      fill_in 'Сколько лет вы работаете в рекламе:', :with => '2'
      fill_in 'Опыт работы', :with => 'testicko'
      fill_in 'Ваш комментарий', :with => 'comment'
      click_button 'Отправить'
      page.should have_text('Спасибо! Ваше сообщение отправлено')
    end
  end
end
