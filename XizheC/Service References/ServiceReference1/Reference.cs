﻿//------------------------------------------------------------------------------
// <auto-generated>
//     此代码由工具生成。
//     运行时版本:2.0.50727.8670
//
//     对此文件的更改可能会导致不正确的行为，并且如果
//     重新生成代码，这些更改将会丢失。
// </auto-generated>
//------------------------------------------------------------------------------

namespace XizheC.ServiceReference1 {
    
    
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "3.0.0.0")]
    [System.ServiceModel.ServiceContractAttribute(Namespace="http://oyxxi.com/", ConfigurationName="ServiceReference1.Service1Soap")]
    public interface Service1Soap {
        
        // CODEGEN: 命名空间 http://oyxxi.com/ 的元素名称 HelloWorldResult 以后生成的消息协定未标记为 nillable
        [System.ServiceModel.OperationContractAttribute(Action="http://oyxxi.com/HelloWorld", ReplyAction="*")]
        XizheC.ServiceReference1.HelloWorldResponse HelloWorld(XizheC.ServiceReference1.HelloWorldRequest request);
        
        // CODEGEN: 命名空间 http://oyxxi.com/ 的元素名称 USER_NAME 以后生成的消息协定未标记为 nillable
        [System.ServiceModel.OperationContractAttribute(Action="http://oyxxi.com/getsqlcon", ReplyAction="*")]
        XizheC.ServiceReference1.getsqlconResponse getsqlcon(XizheC.ServiceReference1.getsqlconRequest request);
    }
    
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "3.0.0.0")]
    [System.ServiceModel.MessageContractAttribute(IsWrapped=false)]
    public partial class HelloWorldRequest {
        
        [System.ServiceModel.MessageBodyMemberAttribute(Name="HelloWorld", Namespace="http://oyxxi.com/", Order=0)]
        public XizheC.ServiceReference1.HelloWorldRequestBody Body;
        
        public HelloWorldRequest() {
        }
        
        public HelloWorldRequest(XizheC.ServiceReference1.HelloWorldRequestBody Body) {
            this.Body = Body;
        }
    }
    
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "3.0.0.0")]
    [System.Runtime.Serialization.DataContractAttribute()]
    public partial class HelloWorldRequestBody {
        
        public HelloWorldRequestBody() {
        }
    }
    
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "3.0.0.0")]
    [System.ServiceModel.MessageContractAttribute(IsWrapped=false)]
    public partial class HelloWorldResponse {
        
        [System.ServiceModel.MessageBodyMemberAttribute(Name="HelloWorldResponse", Namespace="http://oyxxi.com/", Order=0)]
        public XizheC.ServiceReference1.HelloWorldResponseBody Body;
        
        public HelloWorldResponse() {
        }
        
        public HelloWorldResponse(XizheC.ServiceReference1.HelloWorldResponseBody Body) {
            this.Body = Body;
        }
    }
    
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "3.0.0.0")]
    [System.Runtime.Serialization.DataContractAttribute(Namespace="http://oyxxi.com/")]
    public partial class HelloWorldResponseBody {
        
        [System.Runtime.Serialization.DataMemberAttribute(EmitDefaultValue=false, Order=0)]
        public string HelloWorldResult;
        
        public HelloWorldResponseBody() {
        }
        
        public HelloWorldResponseBody(string HelloWorldResult) {
            this.HelloWorldResult = HelloWorldResult;
        }
    }
    
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "3.0.0.0")]
    [System.ServiceModel.MessageContractAttribute(IsWrapped=false)]
    public partial class getsqlconRequest {
        
        [System.ServiceModel.MessageBodyMemberAttribute(Name="getsqlcon", Namespace="http://oyxxi.com/", Order=0)]
        public XizheC.ServiceReference1.getsqlconRequestBody Body;
        
        public getsqlconRequest() {
        }
        
        public getsqlconRequest(XizheC.ServiceReference1.getsqlconRequestBody Body) {
            this.Body = Body;
        }
    }
    
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "3.0.0.0")]
    [System.Runtime.Serialization.DataContractAttribute(Namespace="http://oyxxi.com/")]
    public partial class getsqlconRequestBody {
        
        [System.Runtime.Serialization.DataMemberAttribute(EmitDefaultValue=false, Order=0)]
        public string USER_NAME;
        
        [System.Runtime.Serialization.DataMemberAttribute(EmitDefaultValue=false, Order=1)]
        public string PASSWORD;
        
        [System.Runtime.Serialization.DataMemberAttribute(EmitDefaultValue=false, Order=2)]
        public string DOMAIN;
        
        public getsqlconRequestBody() {
        }
        
        public getsqlconRequestBody(string USER_NAME, string PASSWORD, string DOMAIN) {
            this.USER_NAME = USER_NAME;
            this.PASSWORD = PASSWORD;
            this.DOMAIN = DOMAIN;
        }
    }
    
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "3.0.0.0")]
    [System.ServiceModel.MessageContractAttribute(IsWrapped=false)]
    public partial class getsqlconResponse {
        
        [System.ServiceModel.MessageBodyMemberAttribute(Name="getsqlconResponse", Namespace="http://oyxxi.com/", Order=0)]
        public XizheC.ServiceReference1.getsqlconResponseBody Body;
        
        public getsqlconResponse() {
        }
        
        public getsqlconResponse(XizheC.ServiceReference1.getsqlconResponseBody Body) {
            this.Body = Body;
        }
    }
    
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "3.0.0.0")]
    [System.Runtime.Serialization.DataContractAttribute(Namespace="http://oyxxi.com/")]
    public partial class getsqlconResponseBody {
        
        [System.Runtime.Serialization.DataMemberAttribute(EmitDefaultValue=false, Order=0)]
        public string getsqlconResult;
        
        public getsqlconResponseBody() {
        }
        
        public getsqlconResponseBody(string getsqlconResult) {
            this.getsqlconResult = getsqlconResult;
        }
    }
    
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "3.0.0.0")]
    public interface Service1SoapChannel : XizheC.ServiceReference1.Service1Soap, System.ServiceModel.IClientChannel {
    }
    
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "3.0.0.0")]
    public partial class Service1SoapClient : System.ServiceModel.ClientBase<XizheC.ServiceReference1.Service1Soap>, XizheC.ServiceReference1.Service1Soap {
        
        public Service1SoapClient() {
        }
        
        public Service1SoapClient(string endpointConfigurationName) : 
                base(endpointConfigurationName) {
        }
        
        public Service1SoapClient(string endpointConfigurationName, string remoteAddress) : 
                base(endpointConfigurationName, remoteAddress) {
        }
        
        public Service1SoapClient(string endpointConfigurationName, System.ServiceModel.EndpointAddress remoteAddress) : 
                base(endpointConfigurationName, remoteAddress) {
        }
        
        public Service1SoapClient(System.ServiceModel.Channels.Binding binding, System.ServiceModel.EndpointAddress remoteAddress) : 
                base(binding, remoteAddress) {
        }
        
        [System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Advanced)]
        XizheC.ServiceReference1.HelloWorldResponse XizheC.ServiceReference1.Service1Soap.HelloWorld(XizheC.ServiceReference1.HelloWorldRequest request) {
            return base.Channel.HelloWorld(request);
        }
        
        public string HelloWorld() {
            XizheC.ServiceReference1.HelloWorldRequest inValue = new XizheC.ServiceReference1.HelloWorldRequest();
            inValue.Body = new XizheC.ServiceReference1.HelloWorldRequestBody();
            XizheC.ServiceReference1.HelloWorldResponse retVal = ((XizheC.ServiceReference1.Service1Soap)(this)).HelloWorld(inValue);
            return retVal.Body.HelloWorldResult;
        }
        
        [System.ComponentModel.EditorBrowsableAttribute(System.ComponentModel.EditorBrowsableState.Advanced)]
        XizheC.ServiceReference1.getsqlconResponse XizheC.ServiceReference1.Service1Soap.getsqlcon(XizheC.ServiceReference1.getsqlconRequest request) {
            return base.Channel.getsqlcon(request);
        }
        
        public string getsqlcon(string USER_NAME, string PASSWORD, string DOMAIN) {
            XizheC.ServiceReference1.getsqlconRequest inValue = new XizheC.ServiceReference1.getsqlconRequest();
            inValue.Body = new XizheC.ServiceReference1.getsqlconRequestBody();
            inValue.Body.USER_NAME = USER_NAME;
            inValue.Body.PASSWORD = PASSWORD;
            inValue.Body.DOMAIN = DOMAIN;
            XizheC.ServiceReference1.getsqlconResponse retVal = ((XizheC.ServiceReference1.Service1Soap)(this)).getsqlcon(inValue);
            return retVal.Body.getsqlconResult;
        }
    }
}
