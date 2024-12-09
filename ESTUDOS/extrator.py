from bs4 import BeautifulSoup
from openpyxl import Workbook
import pathlib

# Exemplos de XML fornecidos
xml_sisstatus = """
<dataset>
<SisStatus>
<id>17</id>
<descricao>Aguardando CCB/Contrato</descricao>
<final>0</final>
</SisStatus>
<SisStatus>
<id>16</id>
<descricao>Fora do Horário</descricao>
<final>0</final>
</SisStatus>
<SisStatus>
<id>15</id>
<descricao>Dados incompletos</descricao>
<final>0</final>
</SisStatus>
<SisStatus>
<id>14</id>
<descricao>Documentação Incompleta</descricao>
<final>0</final>
</SisStatus>
<SisStatus>
<id>13</id>
<descricao>Devolução</descricao>
<final>0</final>
</SisStatus>
<SisStatus>
<id>10</id>
<descricao>Liberação</descricao>
<final>0</final>
</SisStatus>
<SisStatus>
<id>8</id>
<descricao>Aguardando documentação do Associado ou Cliente</descricao>
<final>0</final>
</SisStatus>
<SisStatus>
<id>7</id>
<descricao>Aguardando Solicitante</descricao>
<final>0</final>
</SisStatus>
<SisStatus>
<id>6</id>
<descricao>Finalizado sem sucesso</descricao>
<final>1</final>
</SisStatus>
<SisStatus>
<id>5</id>
<descricao>Aguardando Associado ou Cliente</descricao>
<final>0</final>
</SisStatus>
<SisStatus>
<id>4</id>
<descricao>Cancelado pelo solicitante</descricao>
<final>1</final>
</SisStatus>
<SisStatus>
<id>3</id>
<descricao>Finalizado com sucesso</descricao>
<final>1</final>
</SisStatus>
<SisStatus>
<id>2</id>
<descricao>Em atendimento</descricao>
<final>0</final>
</SisStatus>
<SisStatus>
<id>1</id>
<descricao>Novo</descricao>
<final>0</final>
</SisStatus>
</dataset>
"""

xml_siscanalatendimento = """
<dataset>
<SisCanalAtendimento>
<id>7</id>
<descricao>Redes Sociais</descricao>
<interno/>
</SisCanalAtendimento>
<SisCanalAtendimento>
<id>6</id>
<descricao>Site</descricao>
<interno/>
</SisCanalAtendimento>
<SisCanalAtendimento>
<id>5</id>
<descricao>Interno</descricao>
<interno>1</interno>
</SisCanalAtendimento>
<SisCanalAtendimento>
<id>4</id>
<descricao>E-mail</descricao>
<interno/>
</SisCanalAtendimento>
<SisCanalAtendimento>
<id>3</id>
<descricao>Whatsapp</descricao>
<interno/>
</SisCanalAtendimento>
<SisCanalAtendimento>
<id>2</id>
<descricao>Ligação telefônica</descricao>
<interno/>
</SisCanalAtendimento>
</dataset>
"""

xml_sisprioridade = """
<dataset>
<SisPrioridade>
<id>5</id>
<descricao>5 - Ampla - 120 horas</descricao>
</SisPrioridade>
<SisPrioridade>
<id>4</id>
<descricao>1 - Alta - 8 horas úteis</descricao>
</SisPrioridade>
<SisPrioridade>
<id>3</id>
<descricao>2 - Média - 16 horas úteis</descricao>
</SisPrioridade>
<SisPrioridade>
<id>2</id>
<descricao>3 - Normal - 24 horas úteis</descricao>
</SisPrioridade>
<SisPrioridade>
<id>1</id>
<descricao>4 - Baixa - 40 horas úteis</descricao>
</SisPrioridade>
</dataset>
"""

xml_sissetor = """
<dataset>
<SisSetor>
<id>1</id>
<descricao>00 - AGÊNCIA SEDE - ATENDIMENTO PF</descricao>
</SisSetor>
<SisSetor>
<id>2</id>
<descricao>00 - AGÊNCIA SEDE - ATENDIMENTO PJ</descricao>
</SisSetor>
<SisSetor>
<id>3</id>
<descricao>01 - AGÊNCIA SP</descricao>
</SisSetor>
<SisSetor>
<id>4</id>
<descricao>02 - AGÊNCIA RJ</descricao>
</SisSetor>
<SisSetor>
<id>5</id>
<descricao>03 - AGÊNCIA MG</descricao>
</SisSetor>
<SisSetor>
<id>6</id>
<descricao>05 - AGÊNCIA TAG SUL</descricao>
</SisSetor>
<SisSetor>
<id>7</id>
<descricao>06 - AGÊNCIA CEI</descricao>
</SisSetor>
<SisSetor>
<id>8</id>
<descricao>07 - AGÊNCIA NB</descricao>
</SisSetor>
<SisSetor>
<id>25</id>
<descricao>08 - AGÊNCIA TAG NORTE</descricao>
</SisSetor>
<SisSetor>
<id>9</id>
<descricao>ARQUIVO</descricao>
</SisSetor>
<SisSetor>
<id>10</id>
<descricao>CADASTRO</descricao>
</SisSetor>
<SisSetor>
<id>23</id>
<descricao>CAIXA</descricao>
</SisSetor>
<SisSetor>
<id>11</id>
<descricao>COBRANÇA ADMINISTRATIVA</descricao>
</SisSetor>
<SisSetor>
<id>12</id>
<descricao>COMITÊ DE CRÉDITO - PF</descricao>
</SisSetor>
<SisSetor>
<id>21</id>
<descricao>COMITÊ DE CRÉDITO - PJ</descricao>
</SisSetor>
<SisSetor>
<id>13</id>
<descricao>CONTABILIDADE</descricao>
</SisSetor>
<SisSetor>
<id>14</id>
<descricao>DEPARTAMENTO PESSOAL</descricao>
</SisSetor>
<SisSetor>
<id>15</id>
<descricao>DIRETORIA</descricao>
</SisSetor>
<SisSetor>
<id>16</id>
<descricao>FINANCEIRO</descricao>
</SisSetor>
<SisSetor>
<id>26</id>
<descricao>FINANCEIRO - DESLIGAMENTO</descricao>
</SisSetor>
<SisSetor>
<id>17</id>
<descricao>FINANCEIRO - LANÇAMENTO (DESATIVADO)</descricao>
</SisSetor>
<SisSetor>
<id>22</id>
<descricao>PRODUTOS</descricao>
</SisSetor>
<SisSetor>
<id>27</id>
<descricao>SETOR DE EMPRÉSTIMO</descricao>
</SisSetor>
<SisSetor>
<id>18</id>
<descricao>TECNOLOGIA</descricao>
</SisSetor>
<SisSetor>
<id>20</id>
<descricao>TELEFONISTA</descricao>
</SisSetor>
<SisSetor>
<id>19</id>
<descricao>TESOURARIA</descricao>
</SisSetor>
</dataset>
"""

xml_systemusers = """
<dataset>
<SystemUsers>
<id>1</id>
<name>Admin</name>
<login>admin</login>
<email>joao.ricardo@skillnet.com.br</email>
<active>y</active>
</SystemUsers>
<SystemUsers>
<id>2</id>
<name>Juliano de Andrade Almeida</name>
<login>julianon4221_00</login>
<email>juliano.andrade@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>3</id>
<name>Antonio Xará Júnior</name>
<login>antoniox4221_00</login>
<email>antonio.xara@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>7</id>
<name>Master</name>
<login>master</login>
<email>tecnologia@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>8</id>
<name>Iury Silva</name>
<login>iurys4221_00</login>
<email>iury.silva@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>9</id>
<name>Saulo Santos</name>
<login>saulov4221_00</login>
<email>saulo.santos@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>10</id>
<name>Ludmilla Silva</name>
<login>ludmillas4221_00</login>
<email>ludmilla.silva@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>11</id>
<name>Renata Ribeiro</name>
<login>renatas4221_00</login>
<email>renata.ribeiro@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>12</id>
<name>Gilmar Lopes</name>
<login>gilmarl4221_00</login>
<email>gilmar.lopes@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>13</id>
<name>Filipe Macedo</name>
<login>filipeb4221_00</login>
<email>filipe.macedo@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>14</id>
<name>Felipe Viana</name>
<login>felipev4221_00</login>
<email>felipe.viana@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>15</id>
<name>Karoline Vaz</name>
<login>karolinep4221_00</login>
<email>karoline.vaz@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>16</id>
<name>Edilma Araújo</name>
<login>edilmaa4221_00</login>
<email>edilma.araujo@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>17</id>
<name>Flavilene Gomes</name>
<login>flavileneg4221_00</login>
<email>flavilene.gomes@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>18</id>
<name>Nathalia Silvestre</name>
<login>nathaliac4221_00</login>
<email>nathalia.silvestre@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>19</id>
<name>Adriana Martins</name>
<login>adrianam4221_00</login>
<email>adriana.martins@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>20</id>
<name>Leticia Oliveira</name>
<login>silvial4221_00</login>
<email>leticia.oliveira@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>21</id>
<name>Saulo Sales</name>
<login>saulor4221_00</login>
<email>saulo.sales@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>22</id>
<name>Rayane Almeida</name>
<login>rayanec4221_00</login>
<email>rayane.almeida@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>23</id>
<name>Lucas Ferreira</name>
<login>lucasf4221_00</login>
<email>lucas.ferreira@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>24</id>
<name>Euclenes Gomes</name>
<login>euclenesg4221_00</login>
<email>euclenes.gomes@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>25</id>
<name>Raphael Silva</name>
<login>raphaelf4221_00</login>
<email>raphael.silva@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>26</id>
<name>Ana Coralina</name>
<login>anav4221_00</login>
<email>ana.coralina@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>27</id>
<name>Maria da Conceição</name>
<login>mariac4221_00</login>
<email>maria.oliveira@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>28</id>
<name>Aline Martiniano</name>
<login>alines4221_00</login>
<email>aline.martiniano@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>29</id>
<name>Paloma Alves</name>
<login>palomaa4221_00</login>
<email>paloma.alves@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>30</id>
<name>Isabelle Mendes</name>
<login>isabellea4221_00</login>
<email>isabelle.mendes@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>31</id>
<name>Maria Docarmo</name>
<login>mariag4221_00</login>
<email>maria.docarmo@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>32</id>
<name>Gustavo Coelho</name>
<login>gustavoj4221_00</login>
<email>gustavo.coelho@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>33</id>
<name>Lidiane Santos</name>
<login>lidianep4221_00</login>
<email>lidiane.santos@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>34</id>
<name>João Luiz</name>
<login>joaol4221_00</login>
<email>joao.luiz@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>35</id>
<name>Fabimael Silva</name>
<login>fabimaels4221_00</login>
<email>fabimael.silva@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>36</id>
<name>Daniela Barbosa</name>
<login>danielac4221_00</login>
<email>daniela.barbosa@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>37</id>
<name>Rubens Assis</name>
<login>rubensa4221_00</login>
<email>rubens.assis@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>38</id>
<name>José Ricardo</name>
<login>josem4221_00</login>
<email>jose.ricardo@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>39</id>
<name>Francisnaldo Batista</name>
<login>francisnaldob4221_00</login>
<email>francisnaldo.batista@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>40</id>
<name>Erika Fernandes</name>
<login>erikam4221_00</login>
<email>erika.fernandes@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>41</id>
<name>Diene Santos</name>
<login>dienes4221_00</login>
<email>diene.santos@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>42</id>
<name>Aloma Petiani</name>
<login>alomap4221_00</login>
<email>aloma.petiani@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>43</id>
<name>Alair Vieira</name>
<login>alairv4221_00</login>
<email>alair.vieira@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>44</id>
<name>Kelly Machado</name>
<login>kellym4221_00</login>
<email>kelly.machado@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>45</id>
<name>Marcelo Brandao</name>
<login>marcelog4221_00</login>
<email>marcelo.brandao@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>46</id>
<name>Eliza Ferreira</name>
<login>elizac4221_00</login>
<email>eliza.ferreira@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>47</id>
<name>Josy Alves</name>
<login>josilaniaa4221_00</login>
<email>josy.alves@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>48</id>
<name>Abadia Iglesias</name>
<login>mariaa4221</login>
<email>abadia.iglesias@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>49</id>
<name>Patricia Pontes</name>
<login>patriciah4221_00</login>
<email>patricia.pontes@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>50</id>
<name>Valéria Souza</name>
<login>valeriac4221_00</login>
<email>valeria.souza@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>51</id>
<name>André Miranda</name>
<login>andrev4221_00</login>
<email>andre.miranda@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>52</id>
<name>Elisregina Medeiros</name>
<login>elisreginam4221_00</login>
<email>elisregina.medeiros@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>53</id>
<name>Enilza Faria</name>
<login>enilzam4221_00</login>
<email>enilza.faria@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>54</id>
<name>Emanuel Araújo</name>
<login>emanuela4221_00</login>
<email>emanuel.araujo@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>55</id>
<name>Marcelindo Braga</name>
<login>marcelindob4221_00</login>
<email>marcelindo.braga@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>56</id>
<name>Alcindo Pironi</name>
<login>alcindop4221_00</login>
<email>alcindo.pironi@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>57</id>
<name>Rutemberg Cesar</name>
<login>rutembergc4221_00</login>
<email>rutemberg.cesar@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>58</id>
<name>Lidia Souza</name>
<login>lidias4221_00</login>
<email>lidia.souza@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>59</id>
<name>Valdir Santos</name>
<login>valdird4221_00</login>
<email>valdir.santos@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>60</id>
<name>Alessandro Alves</name>
<login>alessandroa4221_00</login>
<email>alessandro.alves@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>61</id>
<name>Gardenia Ramos</name>
<login>gardeniar4221_00</login>
<email>gardenia.freitas@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>62</id>
<name>Thiago Belém</name>
<login>thiagos4221_00</login>
<email>thiago.belem@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>63</id>
<name>Ronan Alves</name>
<login>ronana4221_00</login>
<email>ronan.alves@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>64</id>
<name>Andressa Leite</name>
<login>andressal4221_00</login>
<email>andressa.leite@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>65</id>
<name>Luan Gomes</name>
<login>luanj4221_00</login>
<email>luan.jesus@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>66</id>
<name>Leonardo Mello</name>
<login>leonardoa4221_00</login>
<email>leonardo.mello@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>67</id>
<name>Emanuele Gomes</name>
<login>emanuelep4221_00</login>
<email>emanuele.gomes@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>68</id>
<name>Paulo Sousa</name>
<login>pauloh4221_00</login>
<email>paulo.sousa@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>69</id>
<name>Ludmila Costa</name>
<login>ludmilac4221_00</login>
<email>ludmila.costa@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>70</id>
<name>Juliana Oliveira</name>
<login>julianag4221_00</login>
<email>juliana.oliveira@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>71</id>
<name>Renato Santos</name>
<login>renatoh4221_00</login>
<email>renato.santos@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>72</id>
<name>Joselina Santos</name>
<login>joselinar4221_00</login>
<email>joselina.santos@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>73</id>
<name>Gustavo Santana</name>
<login>gustavor4221_00</login>
<email>gustavo.santana@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>74</id>
<name>Jaqueline Santos</name>
<login>jaquelinea4221_00</login>
<email>jaqueline.santos@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>75</id>
<name>Elidiane Oliveira</name>
<login>elidianel4221_00</login>
<email>elidiane.oliveira@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>76</id>
<name>Daniela Borges</name>
<login>danielaa4221_00</login>
<email>daniela.borges@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>77</id>
<name>Paulo Rodrigues</name>
<login>paulos4221_00</login>
<email>paulo.rodrigues@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>78</id>
<name>Nelson Pessuto</name>
<login>nelsonp4221_00</login>
<email>nelson.pessuto@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>79</id>
<name>Carlos Pio</name>
<login>carlosp4221_00</login>
<email>carlos.pio@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>80</id>
<name>Raiàne Carvâlho</name>
<login>raianes4221_00</login>
<email>raiane.santos@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>81</id>
<name>Alexandre Teles</name>
<login>alexandret4221_00</login>
<email>alexandre.teles@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>82</id>
<name>Bruno Cesar</name>
<login>brunoc4221_00</login>
<email>bruno.cesar@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>83</id>
<name>Leonardo Linhares</name>
<login>leonardol4221_00</login>
<email>leonardo.linhares@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>84</id>
<name>Mônica Cordeiro</name>
<login>MonicaA4221_00</login>
<email>monica.cordeiro@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>85</id>
<name>Keila Cruz</name>
<login>keilac4221_00</login>
<email>keila.cruz@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>86</id>
<name>Pedro Santos</name>
<login>PedroH4221_00</login>
<email>pedro.santos@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>87</id>
<name>teste</name>
<login>usuariot4221_00</login>
<email>teste</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>88</id>
<name>Vinícius de Carvalho Dias</name>
<login>viniciusc4221_00</login>
<email>vinicius.dias@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>89</id>
<name>Joel Lima de Sousa Júnior</name>
<login>joell4221_00</login>
<email>joel.lima@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>90</id>
<name>Eimart Hebert Freitas Rocha</name>
<login>EimartH4221_00</login>
<email>eimart.freitas@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>91</id>
<name>Dayse Pereira da Silva</name>
<login>dayseP4221_00</login>
<email>dayse.silva@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>92</id>
<name>Larissa Natália Silva de Andrade Sousa</name>
<login>LarissaN4221_00</login>
<email>larissa.sousa@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>93</id>
<name>Bruna Aguiar</name>
<login>brunav4221_00</login>
<email>bruna.aguiar@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>94</id>
<name>Simone Ferreira</name>
<login>simonea4221_00</login>
<email>simone.ferreira@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>95</id>
<name>Thiago Mendes</name>
<login>thiagoc4221_00</login>
<email>thiago.mendes@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>96</id>
<name>Ester Maia</name>
<login>AlziristerF4221_00</login>
<email>ester.maia@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>97</id>
<name>Luiz Lara</name>
<login>LuizF4221_00</login>
<email>luiza.lara@credfaz.org.be</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>98</id>
<name>Alex Batista dos Santos</name>
<login>AlexB4221_00</login>
<email>alex.batista@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>99</id>
<name>Ana Santos</name>
<login>AnaO4221_00</login>
<email>ana.santos@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>100</id>
<name>Keslon Dias</name>
<login>KeslonM4221_00</login>
<email>keslon.dias</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>101</id>
<name>Mônica Raquel Nunes Carvalho</name>
<login>MonicaR4221_00</login>
<email>monica.carvalho@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>102</id>
<name>Vanja Marrocos</name>
<login>VanjaA4221_00</login>
<email>vanja.marrocos@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>103</id>
<name>Josyane Alves</name>
<login>JosyaneA4221_00</login>
<email>josyane.aguiar@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>104</id>
<name>Fernanda Santos</name>
<login>FernandaP4221_00</login>
<email>fernanda.santos@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>105</id>
<name>Billy de Amorim Santos</name>
<login>billya4221_00</login>
<email>billy.santos@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>106</id>
<name>Erica Peniche</name>
<login>EricaC4221_00</login>
<email>erica.martins@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>107</id>
<name>Alessandro Brasil</name>
<login>AlessandroP4221_00 </login>
<email>alessandro.brasil</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>108</id>
<name>Marcos Seixas de Brito</name>
<login>marcosS4221_00</login>
<email>marcos.brito@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>109</id>
<name>Letícia Dias</name>
<login>leticiab4221_00</login>
<email>leticia.dias@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>110</id>
<name>Maria Gabrielle</name>
<login>MariaB4221_00</login>
<email>gabrielle.magalhaes@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>111</id>
<name>Bruno Rodrigues</name>
<login>brunoa4221_00</login>
<email>bruno.rodrigues@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>112</id>
<name>Bruna Loraine Nogueira Barreto</name>
<login>BrunaL4221_00</login>
<email>bruna.barreto@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>113</id>
<name>Flávia Guimarães</name>
<login>FlaviaS4221_00</login>
<email>flavia.guimaraes@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>114</id>
<name>Alessandra Ribeiro Costa</name>
<login>alessandrar4221_00</login>
<email>alessandra.costa@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>115</id>
<name>Rafael do Nascimento Oliveira</name>
<login>rafaeln4221_00</login>
<email>rafael.oliveira@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>116</id>
<name>Leticia Almeida Barbosa</name>
<login>leticiaa4221_00</login>
<email>leticia.almeida@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>117</id>
<name>Camila Campos Silva</name>
<login>camilac4221_00</login>
<email>camila.campos@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>118</id>
<name>Altamiro Xavier Toledo Junior</name>
<login>altamirox4221_00</login>
<email>altamiro.junior@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>119</id>
<name>Francisca Santos</name>
<login>FranciscaK4221_00</login>
<email>francisca.santos@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>120</id>
<name>Giselle Lima</name>
<login>GiselleS4221_00</login>
<email>giselle.lima@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>121</id>
<name> Adrianna Santos</name>
<login>adriannac4221_00</login>
<email>adrianna.santos@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>122</id>
<name>Karoline Mendes</name>
<login>karolinem4221_00</login>
<email>karoline.mendes@credfaz.org.br</email>
<active>N</active>
</SystemUsers>
<SystemUsers>
<id>123</id>
<name>Matheus Ribeiro</name>
<login>MatheusB4221_00</login>
<email>matheus.ribeiro@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
<SystemUsers>
<id>124</id>
<name>Juraci Schwantes</name>
<login>JuraciM4221_00</login>
<email>juraci.schwantes@credfaz.org.br</email>
<active>Y</active>
</SystemUsers>
</dataset>
"""

xml_siscategoria = """
<dataset>
<SisCategoria>
<id>123</id>
<descricao>CONTA CORRENTE - ACESSO AO APP</descricao>
</SisCategoria>
<SisCategoria>
<id>122</id>
<descricao>CAPITAL - SALDO DE CONTA CAPITAL</descricao>
</SisCategoria>
<SisCategoria>
<id>121</id>
<descricao>PROTOCOLO</descricao>
</SisCategoria>
<SisCategoria>
<id>120</id>
<descricao>OPERAÇÃO DE CRÉDITO - SOLICITAÇÃO DE CÓPIA DE CONTRATO</descricao>
</SisCategoria>
<SisCategoria>
<id>119</id>
<descricao>SIPAG - SOLICITAÇÃO DE MAQUININHA</descricao>
</SisCategoria>
<SisCategoria>
<id>118</id>
<descricao>CONCILIAÇÃO CONTA CONVÊNIO</descricao>
</SisCategoria>
<SisCategoria>
<id>117</id>
<descricao>CAMPANHA DE MARKETING</descricao>
</SisCategoria>
<SisCategoria>
<id>116</id>
<descricao>ASSOCIAÇÃO OU ABERTURA DE CONTA DIGITAL</descricao>
</SisCategoria>
<SisCategoria>
<id>115</id>
<descricao>OPERAÇÃO DE CRÉDITO - CONSIGNADO BANCOOB</descricao>
</SisCategoria>
<SisCategoria>
<id>114</id>
<descricao>OPERAÇÃO DE CRÉDITO - CONSIGNADO INSS</descricao>
</SisCategoria>
<SisCategoria>
<id>113</id>
<descricao>CARTÃO - DESBLOQUEIO DE LIMITE DE CRÉDITO</descricao>
</SisCategoria>
<SisCategoria>
<id>112</id>
<descricao>OPERAÇÃO DE CRÉDITO - FINANCIAMENTO DE ENERGIA FOTOVOLTAICA</descricao>
</SisCategoria>
<SisCategoria>
<id>111</id>
<descricao>CARTÃO - PARCELAMENTO DE FATURA</descricao>
</SisCategoria>
<SisCategoria>
<id>109</id>
<descricao>OPERAÇÃO DE CRÉDITO - SOLICITAÇÃO DE DED</descricao>
</SisCategoria>
<SisCategoria>
<id>108</id>
<descricao>DEVOLUÇÃO - OUTRAS</descricao>
</SisCategoria>
<SisCategoria>
<id>107</id>
<descricao>DEVOLUÇÃO - SEGURO PRESTAMISTA</descricao>
</SisCategoria>
<SisCategoria>
<id>106</id>
<descricao>DEVOLUÇÃO - PARCELA DE EMPRÉSTIMO</descricao>
</SisCategoria>
<SisCategoria>
<id>105</id>
<descricao>CARTA FIANÇA</descricao>
</SisCategoria>
<SisCategoria>
<id>104</id>
<descricao>OPERAÇÃO DE CRÉDITO - CRÉDITO BNDES </descricao>
</SisCategoria>
<SisCategoria>
<id>103</id>
<descricao>OPERAÇÃO DE CRÉDITO - FINANCIAMENTO DE VEÍCULOS</descricao>
</SisCategoria>
<SisCategoria>
<id>102</id>
<descricao>AÇÃO JUDICIAL</descricao>
</SisCategoria>
<SisCategoria>
<id>101</id>
<descricao>PROTESTO</descricao>
</SisCategoria>
<SisCategoria>
<id>100</id>
<descricao>CONFERÊNCIA</descricao>
</SisCategoria>
<SisCategoria>
<id>99</id>
<descricao>COMITÊ DE CRÉDITO</descricao>
</SisCategoria>
<SisCategoria>
<id>97</id>
<descricao>CARTÃO - DESBLOQUEIO</descricao>
</SisCategoria>
<SisCategoria>
<id>96</id>
<descricao>VISITA/PROSPECÇÃO</descricao>
</SisCategoria>
<SisCategoria>
<id>94</id>
<descricao>RECLAMAÇÃO</descricao>
</SisCategoria>
<SisCategoria>
<id>93</id>
<descricao>PAGAMENTO DE PORTABILIDADE</descricao>
</SisCategoria>
<SisCategoria>
<id>92</id>
<descricao>RELATÓRIO DE VISITA</descricao>
</SisCategoria>
<SisCategoria>
<id>91</id>
<descricao>RD STATION</descricao>
</SisCategoria>
<SisCategoria>
<id>90</id>
<descricao>CAPITAL - BAIXA TOTAL/PARCIAL</descricao>
</SisCategoria>
<SisCategoria>
<id>89</id>
<descricao>LANÇAMENTO DE TARIFA</descricao>
</SisCategoria>
<SisCategoria>
<id>88</id>
<descricao>ESTORNO</descricao>
</SisCategoria>
<SisCategoria>
<id>87</id>
<descricao>PROVA DE VIDA</descricao>
</SisCategoria>
<SisCategoria>
<id>86</id>
<descricao>AVISO DE LANÇAMENTO - ACORDO</descricao>
</SisCategoria>
<SisCategoria>
<id>85</id>
<descricao>SOLICITAÇÃO DE APROVAÇÃO</descricao>
</SisCategoria>
<SisCategoria>
<id>84</id>
<descricao>OPERAÇÃO DE CRÉDITO - CANCELAMENTO DE LIMITE DE CHEQUE ESPECIAL - PF</descricao>
</SisCategoria>
<SisCategoria>
<id>83</id>
<descricao>OPERAÇÃO DE CRÉDITO - CANCELAMENTO DE LIMITE DE CONTA GARANTIDA - PJ</descricao>
</SisCategoria>
<SisCategoria>
<id>82</id>
<descricao>DOCUMENTAÇÃO - PROTESTO</descricao>
</SisCategoria>
<SisCategoria>
<id>81</id>
<descricao>DOCUMENTAÇÃO - AÇÃO JUDICIAL</descricao>
</SisCategoria>
<SisCategoria>
<id>80</id>
<descricao>SISBR</descricao>
</SisCategoria>
<SisCategoria>
<id>79</id>
<descricao>CHEQUE - REAPRESENTAÇÃO</descricao>
</SisCategoria>
<SisCategoria>
<id>78</id>
<descricao>CONTA CORRENTE - ENCERRAMENTO</descricao>
</SisCategoria>
<SisCategoria>
<id>77</id>
<descricao>CONFERÊNCIA DE DEPÓSITO</descricao>
</SisCategoria>
<SisCategoria>
<id>76</id>
<descricao>CAPITAL - ABERTURA</descricao>
</SisCategoria>
<SisCategoria>
<id>75</id>
<descricao>LIBERAÇÃO DE ACESSO</descricao>
</SisCategoria>
<SisCategoria>
<id>74</id>
<descricao>CHEQUE - CUSTÓDIA</descricao>
</SisCategoria>
<SisCategoria>
<id>73</id>
<descricao>ATUALIZAÇÃO CADASTRAL</descricao>
</SisCategoria>
<SisCategoria>
<id>72</id>
<descricao>LIMITE CRL</descricao>
</SisCategoria>
<SisCategoria>
<id>71</id>
<descricao>CARTÃO - ADICIONAL</descricao>
</SisCategoria>
<SisCategoria>
<id>70</id>
<descricao>COMUNICAÇÃO E TRATATIVAS DE ÓBITO</descricao>
</SisCategoria>
<SisCategoria>
<id>69</id>
<descricao>SOLICITAÇÃO DE BOLETO</descricao>
</SisCategoria>
<SisCategoria>
<id>68</id>
<descricao>SOLICITAÇÃO DE QUITAÇÃO</descricao>
</SisCategoria>
<SisCategoria>
<id>67</id>
<descricao>SISTEMA DE OS</descricao>
</SisCategoria>
<SisCategoria>
<id>66</id>
<descricao>HARDWARE/PERIFÉRICOS</descricao>
</SisCategoria>
<SisCategoria>
<id>65</id>
<descricao>CARTÃO - IMPLANTAÇÃO DE LIMITE DE CRÉDITO</descricao>
</SisCategoria>
<SisCategoria>
<id>64</id>
<descricao>CARTÃO - BLOQUEIO</descricao>
</SisCategoria>
<SisCategoria>
<id>63</id>
<descricao>CARTEIRA DE CLIENTES</descricao>
</SisCategoria>
<SisCategoria>
<id>62</id>
<descricao>E-MAIL</descricao>
</SisCategoria>
<SisCategoria>
<id>61</id>
<descricao>IMPRESSORA</descricao>
</SisCategoria>
<SisCategoria>
<id>60</id>
<descricao>TESOUREIRO ELETRÔNICO - GLORY</descricao>
</SisCategoria>
<SisCategoria>
<id>59</id>
<descricao>CAIXA</descricao>
</SisCategoria>
<SisCategoria>
<id>58</id>
<descricao>CITRIX</descricao>
</SisCategoria>
<SisCategoria>
<id>57</id>
<descricao>ACESSOS (SISBR/SIPAGNET/REDE/ETC)</descricao>
</SisCategoria>
<SisCategoria>
<id>56</id>
<descricao>OUTROS</descricao>
</SisCategoria>
<SisCategoria>
<id>55</id>
<descricao>PROPOSTA DE CRÉDITO</descricao>
</SisCategoria>
<SisCategoria>
<id>44</id>
<descricao>TÍTULO DE CAPITALIZAÇÃO</descricao>
</SisCategoria>
<SisCategoria>
<id>43</id>
<descricao>SOLICITAÇÃO DE TRANSFERÊNCIA</descricao>
</SisCategoria>
<SisCategoria>
<id>42</id>
<descricao>SIPAG</descricao>
</SisCategoria>
<SisCategoria>
<id>41</id>
<descricao>SICOOB NET - SENHA DE EFETIVAÇÃO</descricao>
</SisCategoria>
<SisCategoria>
<id>40</id>
<descricao>SICOOB NET - LIBERAÇÃO DE DISPOSITIVOS</descricao>
</SisCategoria>
<SisCategoria>
<id>39</id>
<descricao>SICOOB NET - LIBERAÇÃO DE COMPUTADOR</descricao>
</SisCategoria>
<SisCategoria>
<id>38</id>
<descricao>ALTERAÇÃO DE LIMITE DIÁRIO</descricao>
</SisCategoria>
<SisCategoria>
<id>37</id>
<descricao>SEGUROS</descricao>
</SisCategoria>
<SisCategoria>
<id>36</id>
<descricao>RATEIO DE SOBRAS</descricao>
</SisCategoria>
<SisCategoria>
<id>35</id>
<descricao>PREVIDÊNCIA</descricao>
</SisCategoria>
<SisCategoria>
<id>34</id>
<descricao>OUVIDORIA</descricao>
</SisCategoria>
<SisCategoria>
<id>33</id>
<descricao>OPERAÇÃO DE CRÉDITO - PORTABILIDADE (VENDA)</descricao>
</SisCategoria>
<SisCategoria>
<id>32</id>
<descricao>OPERAÇÃO DE CRÉDITO - PORTABILIDADE (COMPRA)</descricao>
</SisCategoria>
<SisCategoria>
<id>31</id>
<descricao>OPERAÇÃO DE CRÉDITO - LIMITE DE CONTA GARANTIDA - PJ</descricao>
</SisCategoria>
<SisCategoria>
<id>30</id>
<descricao>OPERAÇÃO DE CRÉDITO - LIMITE DE CHEQUE ESPECIAL - PF</descricao>
</SisCategoria>
<SisCategoria>
<id>29</id>
<descricao>OPERAÇÃO DE CRÉDITO - CRÉDITO PESSOAL</descricao>
</SisCategoria>
<SisCategoria>
<id>28</id>
<descricao>OPERAÇÃO DE CRÉDITO - CONSIGNADO SERVIDOR PÚBLICO</descricao>
</SisCategoria>
<SisCategoria>
<id>27</id>
<descricao>OPERAÇÃO DE CRÉDITO - CONSIGNADO PRIVADO</descricao>
</SisCategoria>
<SisCategoria>
<id>26</id>
<descricao>OPERAÇÃO DE CRÉDITO - CAPITAL DE GIRO</descricao>
</SisCategoria>
<SisCategoria>
<id>25</id>
<descricao>OPERAÇÃO DE CRÉDITO - ANTECIPAÇÃO DE RECEBÍVEIS</descricao>
</SisCategoria>
<SisCategoria>
<id>24</id>
<descricao>INVESTIMENTOS - RESGATE DE APLICAÇÃO</descricao>
</SisCategoria>
<SisCategoria>
<id>23</id>
<descricao>INVESTIMENTOS - DÚVIDAS GERAIS</descricao>
</SisCategoria>
<SisCategoria>
<id>22</id>
<descricao>INVESTIMENTOS - APLICAÇÃO FINANCEIRA</descricao>
</SisCategoria>
<SisCategoria>
<id>21</id>
<descricao>INFORME DE RENDIMENTOS</descricao>
</SisCategoria>
<SisCategoria>
<id>20</id>
<descricao>INDICAÇÃO / PROSPECÇÃO / LEADS</descricao>
</SisCategoria>
<SisCategoria>
<id>19</id>
<descricao>CONTA CORRENTE - PORTABILIDADE DE SALÁRIO</descricao>
</SisCategoria>
<SisCategoria>
<id>18</id>
<descricao>CONTA CORRENTE</descricao>
</SisCategoria>
<SisCategoria>
<id>17</id>
<descricao>CONSÓRCIOS</descricao>
</SisCategoria>
<SisCategoria>
<id>16</id>
<descricao>COBRANÇA BANCÁRIA</descricao>
</SisCategoria>
<SisCategoria>
<id>15</id>
<descricao>COBRANÇA ADMINISTRATIVA</descricao>
</SisCategoria>
<SisCategoria>
<id>14</id>
<descricao>CARTÃO - SOLICITAÇÃO DE NOVO CARTÃO</descricao>
</SisCategoria>
<SisCategoria>
<id>13</id>
<descricao>CARTÃO - SOLICITAÇÃO DE FATURA</descricao>
</SisCategoria>
<SisCategoria>
<id>12</id>
<descricao>CARTÃO - PRÉ-PAGO</descricao>
</SisCategoria>
<SisCategoria>
<id>11</id>
<descricao>CARTÃO - CONTESTAÇÃO</descricao>
</SisCategoria>
<SisCategoria>
<id>10</id>
<descricao>CARTÃO - CANCELAMENTO</descricao>
</SisCategoria>
<SisCategoria>
<id>9</id>
<descricao>CARTÃO - ALTERAÇÃO DE LIMITE</descricao>
</SisCategoria>
<SisCategoria>
<id>8</id>
<descricao>CAPITAL - RESGATE</descricao>
</SisCategoria>
<SisCategoria>
<id>7</id>
<descricao>CAPITAL - REDUÇÃO DE INTEGRALIZAÇÃO</descricao>
</SisCategoria>
<SisCategoria>
<id>6</id>
<descricao>CAPITAL - DESLIGAMENTO</descricao>
</SisCategoria>
<SisCategoria>
<id>5</id>
<descricao>CAPITAL - CANCELAMENTO DE INTEGRALIZAÇÃO</descricao>
</SisCategoria>
<SisCategoria>
<id>4</id>
<descricao>CAPITAL - AUMENTO DE INTEGRALIZAÇÃO</descricao>
</SisCategoria>
<SisCategoria>
<id>3</id>
<descricao>ASSOCIADO PJ BAIXADO</descricao>
</SisCategoria>
<SisCategoria>
<id>2</id>
<descricao>ASSOCIADO PF FALECIDO</descricao>
</SisCategoria>
<SisCategoria>
<id>1</id>
<descricao>ASSOCIAÇÃO OU ABERTURA DE CONTA</descricao>
</SisCategoria>
</dataset>
"""


# Função para adicionar dados à planilha
def adicionar_dados_ao_excel(xml_data, aba_nome, colunas):
    soup = BeautifulSoup(xml_data, "html.parser")
    ws = wb.create_sheet(title=aba_nome)
    ws.append(colunas)
    for item in soup.find_all(aba_nome.lower()):
        linha = [
            item.find(coluna.lower()).text if item.find(coluna.lower()) else ""
            for coluna in colunas
        ]
        ws.append(linha)


# Criar um novo workbook do Excel
wb = Workbook()
wb.remove(wb.active)  # Remover a aba padrão criada

# Adicionar dados para SisStatus
adicionar_dados_ao_excel(xml_sisstatus, "SisStatus", ["ID", "DESCRICAO", "FINAL"])

# Adicionar dados para SisCanalAtendimento
adicionar_dados_ao_excel(
    xml_siscanalatendimento, "SisCanalAtendimento", ["ID", "DESCRICAO", "INTERNO"]
)

# Adicionar dados para SisPrioridade
adicionar_dados_ao_excel(xml_sisprioridade, "SisPrioridade", ["ID", "DESCRICAO"])

# Adicionar dados para SisSetor
adicionar_dados_ao_excel(xml_sissetor, "SisSetor", ["ID", "DESCRICAO"])

# Adicionar dados para SisCategoria
adicionar_dados_ao_excel(xml_siscategoria, "Siscategoria", ["ID", "DESCRICAO"])

# Adicionar dados para SystemUsers
adicionar_dados_ao_excel(
    xml_systemusers, "Systemusers", ["ID", "NAME", "LOGIN", "EMAIL", "ACTIVE"]
)

# Usar pathlib para definir o caminho do arquivo
output_path = pathlib.Path.home() / "Desktop" / "output_unificado.xlsx"

# Salvar o arquivo Excel
wb.save(output_path)

print(f"Dados extraídos e salvos em {output_path}")
