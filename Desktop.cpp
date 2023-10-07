#include <Windows.h>
#include <ShObjIdl_core.h>
#include <string>
#include <iostream>
#include <fstream>
#pragma comment(lib,"ole32.lib")

using namespace std;

class cFuncoes
{
private:

	IDesktopWallpaper* PapelDeParede = 0;

public:

	void InicializarInstancia()
	{
		CoInitializeEx(0, COINIT_MULTITHREADED);
		CoCreateInstance(CLSID_DesktopWallpaper, 0, CLSCTX_ALL, IID_IDesktopWallpaper, (void**)&PapelDeParede);
	}

	void HabilitarPlanoDeFundo(BOOL Habilitar)
	{
		HRESULT Res;
		Res = PapelDeParede->Enable(Habilitar);
	}

	void ProximoPapelDeParede(bool Proxima)
	{
		HRESULT Res;

		//Esta configura��o se aplica a mudan�a de imagem para o monitor padr�o configurado.
		//Caso seja necess�rio usar em outro monitor, chame GetMonitorDevicePathAt.
		if (Proxima == true)
		{
			Res = PapelDeParede->AdvanceSlideshow(0, DSD_FORWARD);//Pr�xima
		}
		else
		{
			Res = PapelDeParede->AdvanceSlideshow(0, DSD_BACKWARD);//Anterior
		}
	}

	//Para alterar as cores, s�o usados valores hexadecimais.
	//Poder� tamb�m customizar de sua prefer�ncia, efetuando a altera��o nestes valores ou criando-os.
	void AlterarCorDeFundo(COLORREF CorHexadecimal)
	{
		//Exemplos:
		COLORREF Vermelho = 0x000000ff;
		COLORREF Verde = 0x0000ff00;
		COLORREF Azul = 0x00ff0000;
		COLORREF Preto = 0x00000000;

		HRESULT Res;
		//Para que ocorra sucesso na opera��o, n�o pode haver nenhum papel de parede ou o plano de fundo estar habilitado.
		Res = PapelDeParede->SetBackgroundColor(CorHexadecimal);
	}

	COLORREF CorObtida;
	void ObterCorDeFundo()
	{
		HRESULT Res;
		Res = PapelDeParede->GetBackgroundColor(&CorObtida);

		cout << "Cor de fundo: " << CorObtida << '\n';
	}

	/*
	* 
	* Valores e significados para o redimensionamento.
	* 
	* - DWPOS_CENTER - Centralizar.
	* - DWPOS_TILE - Lado a lado.
	* - DWPOS_STRETCH - Preenchimento m�ximo, pode ocorrer perda na qualidade.
	* - DWPOS_FIT - Esticar.
	* - DWPOS_FILL - Preenchimento inteligente.
	* - DWPOS_SPAN - Definir uma �nica imagem em todos os monitores.
	* 
	*/
	void AlterarRedimensionamentoDePapelDeParede()
	{
		HRESULT Res;
		Res = PapelDeParede->SetPosition(DWPOS_FILL);
	}

	void AlterarConfiguracoesDeApresentacaoDeSlides(int Tipo, UINT Tempo)
	{
		HRESULT Res;
		Res = PapelDeParede->SetSlideshowOptions((DESKTOP_SLIDESHOW_OPTIONS)Tipo, Tempo);
	}

	void ObterConfiguracoesDeApresentacaoDeSlides()
	{
		HRESULT Res;

		int Tipo;
		UINT Tempo;//Tempo de mudan�a para a pr�xima imagem.
		Res = PapelDeParede->GetSlideshowOptions((DESKTOP_SLIDESHOW_OPTIONS*)&Tipo,
			&Tempo);

		if (Tipo == 0)
			cout << "A altera��o para imagens aleat�rias est� desativado." << '\n';
		else
			cout << "A altera��o para imagens aleat�rias est� ativo." << '\n';

		cout << "Tempo para a pr�xima mudan�a de imagem: " << Tempo << " millissegundos.\n";
	}

}Funcoes;

int main()
{
	cout << "O assistente est� efetuando configura��es no tema padr�o do Windows...";

	Funcoes.InicializarInstancia();
	Funcoes.ProximoPapelDeParede(true);
	Funcoes.HabilitarPlanoDeFundo(FALSE);
	Funcoes.AlterarCorDeFundo(0x000000ff);
	Funcoes.ObterCorDeFundo();
	Funcoes.HabilitarPlanoDeFundo(TRUE);

	//DSO_SHUFFLEIMAGES para ativar a exibi��o de imagens aleat�rias.
	//60000 mil milissegundos equivalem a 1 minuto para a pr�xima exibi��o.
	Funcoes.AlterarConfiguracoesDeApresentacaoDeSlides(DSO_SHUFFLEIMAGES, 60000);
	Funcoes.ObterConfiguracoesDeApresentacaoDeSlides();

	system("pause");
}