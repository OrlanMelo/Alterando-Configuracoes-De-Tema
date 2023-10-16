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

		//Esta configuração se aplica a mudança de imagem para o monitor padrão configurado.
		//Caso seja necessário usar em outro monitor, chame GetMonitorDevicePathAt.
		if (Proxima == true)
		{
			Res = PapelDeParede->AdvanceSlideshow(0, DSD_FORWARD);//Próxima
		}
		else
		{
			Res = PapelDeParede->AdvanceSlideshow(0, DSD_BACKWARD);//Anterior
		}
	}

	//Para alterar as cores, são usados valores hexadecimais.
	//Poderá também customizar de sua preferência, efetuando a alteração nestes valores ou criando-os.
	void AlterarCorDeFundo(COLORREF CorHexadecimal)
	{
		//Exemplos:
		COLORREF Vermelho = 0x000000ff;
		COLORREF Verde = 0x0000ff00;
		COLORREF Azul = 0x00ff0000;
		COLORREF Preto = 0x00000000;

		HRESULT Res;
		//Para que ocorra sucesso na operação, não pode haver nenhum papel de parede ou o plano de fundo estar habilitado.
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
	* - DWPOS_STRETCH - Preenchimento máximo, pode ocorrer perda na qualidade.
	* - DWPOS_FIT - Esticar.
	* - DWPOS_FILL - Preenchimento inteligente.
	* - DWPOS_SPAN - Definir uma única imagem em todos os monitores.
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
		UINT Tempo;//Tempo de mudança para a próxima imagem.
		Res = PapelDeParede->GetSlideshowOptions((DESKTOP_SLIDESHOW_OPTIONS*)&Tipo,
			&Tempo);

		if (Tipo == 0)
		{
			cout << "A alteração para imagens aleatórias está desativado." << '\n';
		}
		else
		{
			cout << "A alteração para imagens aleatórias está ativo." << '\n';
		}

		cout << "Tempo para a próxima mudança de imagem: " << Tempo << " millissegundos.\n";
	}

}Funcoes;

int main()
{
	cout << "O assistente está efetuando configurações no tema padrão do Windows...";

	Funcoes.InicializarInstancia();
	Funcoes.ProximoPapelDeParede(true);
	Funcoes.HabilitarPlanoDeFundo(FALSE);
	Funcoes.AlterarCorDeFundo(0x000000ff);
	Funcoes.ObterCorDeFundo();
	Funcoes.HabilitarPlanoDeFundo(TRUE);

	//DSO_SHUFFLEIMAGES para ativar a exibição de imagens aleatórias.
	//60000 mil milissegundos equivalem a 1 minuto para a próxima exibição.
	Funcoes.AlterarConfiguracoesDeApresentacaoDeSlides(DSO_SHUFFLEIMAGES, 60000);
	Funcoes.ObterConfiguracoesDeApresentacaoDeSlides();

	system("pause");
}
