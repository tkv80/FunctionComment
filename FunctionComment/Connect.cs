using System;
using EnvDTE;
using EnvDTE80;
using Extensibility;
using Microsoft.VisualStudio.CommandBars;

namespace FunctionComment
{
	/// <summary>추가 기능을 구현하는 개체입니다.</summary>
	/// <seealso class='IDTExtensibility2' />
	public class Connect : IDTExtensibility2, IDTCommandTarget
	{
		/// <summary>추가 기능 개체에 대한 생성자를 구현합니다. 이 메서드 안에 초기화 코드를 배치하십시오.</summary>
		public Connect()
		{
		}

		/// <summary>IDTExtensibility2 인터페이스의 OnConnection 메서드를 구현합니다. 추가 기능이 로드되고 있다는 알림 메시지를 받습니다.</summary>
		/// <param term='application'>호스트 응용 프로그램의 루트 개체입니다.</param>
		/// <param term='connectMode'>추가 기능이 로드되는 방법을 설명합니다.</param>
		/// <param term='addInInst'>이 추가 기능을 나타내는 개체입니다.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
		{
			_applicationObject = (DTE2)application;
			_addInInstance = (AddIn)addInInst;
			if(connectMode == ext_ConnectMode.ext_cm_UISetup)
			{
				object []contextGUIDS = new object[] { };
				Commands2 commands = (Commands2)_applicationObject.Commands;
				string toolsMenuName = "Tools";

				//도구 메뉴에 명령을 배치합니다.
				//주 메뉴 항목이 모두 들어 있는 최상위 명령 모음인 MenuBar 명령 모음 찾기:
				Microsoft.VisualStudio.CommandBars.CommandBar menuBarCommandBar = ((Microsoft.VisualStudio.CommandBars.CommandBars)_applicationObject.CommandBars)["MenuBar"];

				//MenuBar 명령 모음에서 도구 명령 모음 찾기:
				CommandBarControl toolsControl = menuBarCommandBar.Controls[toolsMenuName];
				CommandBarPopup toolsPopup = (CommandBarPopup)toolsControl;

				//추가 기능에서 처리할 명령을 여러 개 추가하면 이 try/catch 블록이 중복될 수 있습니다.
				//  이 경우 QueryStatus/Exec 메서드도 업데이트하여 새 명령 이름을 포함하기만 하면 됩니다.
				try
				{
					//명령 컬렉션에 명령 추가:
					Command command = commands.AddNamedCommand2(_addInInstance, "FunctionComment", "FunctionComment", "Executes the command for FunctionComment", true, 59, ref contextGUIDS, (int)vsCommandStatus.vsCommandStatusSupported+(int)vsCommandStatus.vsCommandStatusEnabled, (int)vsCommandStyle.vsCommandStylePictAndText, vsCommandControlType.vsCommandControlTypeButton);

					//도구 메뉴에 명령에 대한 컨트롤 추가:
					if((command != null) && (toolsPopup != null))
					{
						command.AddControl(toolsPopup.CommandBar, 1);
					}
				}
				catch(System.ArgumentException)
				{
					//이 경우, 같은 이름의 명령이 이미 있기 때문에 예외가 발생할 수
					//  있습니다. 이 경우 명령을 다시 만들 필요가 없으며 예외를 무시해도 
                    //  됩니다.
				}
			}
		}

		/// <summary>IDTExtensibility2 인터페이스의 OnDisconnection 메서드를 구현합니다. 추가 기능이 언로드되고 있다는 알림 메시지를 받습니다.</summary>
		/// <param term='disconnectMode'>추가 기능이 언로드되는 방법을 설명합니다.</param>
		/// <param term='custom'>호스트 응용 프로그램과 관련된 매개 변수의 배열입니다.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
		{
		}

		/// <summary>IDTExtensibility2 인터페이스의 OnAddInsUpdate 메서드를 구현합니다. 추가 기능의 컬렉션이 변경되면 알림 메시지를 받습니다.</summary>
		/// <param term='custom'>호스트 응용 프로그램과 관련된 매개 변수의 배열입니다.</param>
		/// <seealso class='IDTExtensibility2' />		
		public void OnAddInsUpdate(ref Array custom)
		{
		}

		/// <summary>IDTExtensibility2 인터페이스의 OnStartupComplete 메서드를 구현합니다. 호스트 응용 프로그램에서 로드가 완료되었다는 알림 메시지를 받습니다.</summary>
		/// <param term='custom'>호스트 응용 프로그램과 관련된 매개 변수의 배열입니다.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnStartupComplete(ref Array custom)
		{
		}

		/// <summary>IDTExtensibility2 인터페이스의 OnBeginShutdown 메서드를 구현합니다. 호스트 응용 프로그램이 언로드되고 있다는 알림 메시지를 받습니다.</summary>
		/// <param term='custom'>호스트 응용 프로그램과 관련된 매개 변수의 배열입니다.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnBeginShutdown(ref Array custom)
		{
		}
		
		/// <summary>IDTCommandTarget 인터페이스의 QueryStatus 메서드를 구현합니다. 이 메서드는 명령의 사용 여부가 업데이트되면 호출됩니다.</summary>
		/// <param term='commandName'>상태를 확인할 명령의 이름입니다.</param>
		/// <param term='neededText'>명령에 필요한 텍스트입니다.</param>
		/// <param term='status'>사용자 인터페이스에서의 명령 상태입니다.</param>
		/// <param term='commandText'>neededText 매개 변수에서 요청한 텍스트입니다.</param>
		/// <seealso class='Exec' />
		public void QueryStatus(string commandName, vsCommandStatusTextWanted neededText, ref vsCommandStatus status, ref object commandText)
		{
			if(neededText == vsCommandStatusTextWanted.vsCommandStatusTextWantedNone)
			{
				if(commandName == "FunctionComment.Connect.FunctionComment")
				{
					status = (vsCommandStatus)vsCommandStatus.vsCommandStatusSupported|vsCommandStatus.vsCommandStatusEnabled;
					return;
				}
			}
		}

		/// <summary>IDTCommandTarget 인터페이스의 Exec 메서드를 구현합니다. 이 메서드는 명령이 호출되면 호출됩니다.</summary>
		/// <param term='commandName'>실행할 명령의 이름입니다.</param>
		/// <param term='executeOption'>명령을 실행하는 방법을 설명합니다.</param>
		/// <param term='varIn'>호출자에서 명령 처리기로 전달된 매개 변수입니다.</param>
		/// <param term='varOut'>명령 처리기에서 호출자로 전달된 매개 변수입니다.</param>
		/// <param term='handled'>명령이 처리되었는지 여부를 호출자에게 알립니다.</param>
		/// <seealso class='Exec' />
		public void Exec(string commandName, vsCommandExecOption executeOption, ref object varIn, ref object varOut, ref bool handled)
		{
			handled = false;
			if(executeOption == vsCommandExecOption.vsCommandExecOptionDoDefault)
			{
				if(commandName == "FunctionComment.Connect.FunctionComment")
				{
                    FunctionComment();
					handled = true;
					return;
				}
			}
		}

        private void FunctionComment()
        {
            
            EnvDTE.TextSelection textSelection = (EnvDTE.TextSelection)_applicationObject.ActiveWindow.Selection;

            textSelection.StartOfDocument();
            
            textSelection.Insert("// *");
            
            textSelection.NewLine();
            textSelection.Insert("// * @author  모바일 기술팀 ");
            textSelection.Insert(System.Environment.UserName);
            
            textSelection.NewLine();
            textSelection.Insert("// * @date    ");
            textSelection.Insert(System.DateTime.Now.ToLongDateString());
            textSelection.Insert(" ");
            textSelection.Insert(System.DateTime.Now.ToLongTimeString());
            
            textSelection.NewLine();
            textSelection.Insert("// *");

            textSelection.NewLine();
            textSelection.NewLine();
        }

		private DTE2 _applicationObject;
		private AddIn _addInInstance;
	}
}