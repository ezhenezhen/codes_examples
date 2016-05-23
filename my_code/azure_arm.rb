# Installation instructions:
# http://dl.bintray.com/oneclick/rubyinstaller/rubyinstaller-2.2.4.exe
# http://dl.bintray.com/oneclick/rubyinstaller/DevKit-mingw64-32-4.7.2-20130224-1151-sfx.exe (don't use x64 since it may cause problems; add to PATH while installing)
# in cmd.exe:
# cd C:\RubyDevKit
# ruby dk.rb init
# ruby dk.rb install

# http://curl.haxx.se/ca/cacert.pem save it somwhere, define a new System Variable with SSL_CERT_FILE=path/to/file
# additional info about this pem file: https://superdevresources.com/ssl-error-ruby-gems-windows/ 
# in cmd.exe:
# gem install azure_mgmt_compute

#######################################################################################################################

# Binx azure credentials information:
# tenant_id       = 'tenant_id'
# secret          = 'secret'
# client_id       = 'client_id'
# subscription_id = 'subscription_id'

#######################################################################################################################
#######################################################################################################################

# Usage: 
# tenant_id, client_id, secret and subscription_id are constants(are configured for each account) which can be configured on azure cloud.
# more info about getting those:
# https://azure.microsoft.com/en-us/documentation/articles/resource-group-authenticate-service-principal/
# https://blogs.msdn.microsoft.com/tomholl/2014/11/24/unattended-authentication-to-azure-management-apis-with-azure-active-directory/
# http://www.dushyantgill.com/blog/2015/05/23/developers-guide-to-auth-with-azure-resource-manager-api/
# https://azure.microsoft.com/en-us/documentation/articles/resource-group-create-service-principal-portal/

#######################################################################################################################

# power_off:
# ruby azure_arm.rb --tenant_id=tenant_id --client_id=client_id --secret=secret --subscription_id=subscription_id --vm_name=test --resource_group=DEFAULT-WEB-WESTUS --action=power_off

#######################################################################################################################

# machine_state:
# ruby azure_arm.rb --tenant_id=tenant_id --client_id=client_id --secret=secret --subscription_id=subscription_id --vm_name=test --resource_group=DEFAULT-WEB-WESTUS --action=machine_state

#######################################################################################################################

# start:
# ruby azure_arm.rb --tenant_id=tenant_id --client_id=client_id --secret=secret --subscription_id=subscription_id --vm_name=test --resource_group=DEFAULT-WEB-WESTUS --action=start

#######################################################################################################################

# get_vm_list:
# ruby azure_arm.rb --tenant_id=tenant_id --client_id=client_id --secret=secret --subscription_id=subscription_id --action=get_vm_list

# each succeeded event should respond with 200(OK)(if its a command like start, power off, restart)
# other commands like machine state or get vm list will respond with JSON
# [{"subscription":"subscription_id","resource_group":"DEFAULT-WEB-WESTUS","name":"test","state":"PowerState/stopped"}]
# or a string 'PowerState/stopped'

require 'optparse'
require 'azure_mgmt_compute'
include Azure::ARM::Compute
include Azure::ARM::Compute::Models

class AzureArm
  def initialize
    @opts = {}
    OptionParser.new do |opt|
      opt.on('--tenant_id TENANTID')             { |o| @opts[:tenant_id]       = o }
      opt.on('--secret SECRET')                  { |o| @opts[:secret]          = o }
      opt.on('--client_id CLIENTID')             { |o| @opts[:client_id]       = o }
      opt.on('--subscription_id SUBSCRIPTIONID') { |o| @opts[:subscription_id] = o }
      opt.on('--resource_group RESOURCEGROUP')   { |o| @opts[:resource_group]  = o }
      opt.on('--vm_name VMNAME')                 { |o| @opts[:vm_name]         = o }
      opt.on('--action ACTION')                  { |o| @opts[:action]          = o }
    end.parse!

    token_provider = MsRestAzure::ApplicationTokenProvider.new(
      @opts[:tenant_id], 
      @opts[:client_id], 
      @opts[:secret]
    )
    credentials             = MsRest::TokenCredentials.new(token_provider)
    @client                 = ComputeManagementClient.new(credentials)
    @client.subscription_id = @opts[:subscription_id]
    @vm_name                = @opts[:vm_name]
    @resource_group         = @opts[:resource_group]
  end

  def do
    case @opts[:action]
    when "get_vm_list"
      get_vm_list
    when "machine_state"
      machine_state
    when "power_off"
      power_off
    when "start"
      start
    when "restart"
      restart
    end
  end

  private

  def get_vm_list
    vm = []
    virtual_machines = @client.virtual_machines.list_all.value!.body.value
    virtual_machines.each do |machine|
      resource_group = machine.id.split('/')[4]
      vm_name        = machine.id.split('/')[8]
      vm << {
        subscription:   machine.id.split('/')[2],
        resource_group: resource_group,
        name:           vm_name,
        state:          get_machine_state(resource_group, vm_name)
      }
    end
    puts vm.to_json
  end

  def get_machine_state(resource_group, vm_name)
    @client.virtual_machines.get(resource_group, vm_name, 'instanceView').value!.body.properties.instance_view.statuses[1].code
  end

  def machine_state
    puts get_machine_state(@resource_group, @vm_name)
  end

  def power_off
    result = @client.virtual_machines.power_off(@resource_group, @vm_name).value!
    puts result.response.status
  end

  def start
    result = @client.virtual_machines.start(@resource_group, @vm_name).value!
    puts result.response.status
  end

  def restart
    result = @client.virtual_machines.restart(@resource_group, @vm_name).value!
    puts result.response.status
  end
end

AzureArm.new.do
