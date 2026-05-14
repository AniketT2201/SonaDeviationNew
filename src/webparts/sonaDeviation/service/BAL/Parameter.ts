import { ISonaDeviationProps } from '../../components/ISonaDeviationProps';
import SPCRUDOPS from '../DAL/spcrudops';
import { IParameterType } from '../INTERFACE/IParameter';

export interface ParameterTypeOps {
    getAllParameterType(props: ISonaDeviationProps): Promise<IParameterType[]>;
    getParameterTypeById(Id: string | number, props: ISonaDeviationProps): Promise<IParameterType>;
    getParameterType(columnsToRetrieve: string, columnsToExpand: string, filters: string,
    orderby: { column: string; isAscending: boolean }, props: ISonaDeviationProps): Promise<IParameterType[]>;
    getTopParameterType(columnsToRetrieve: string, columnsToExpand: string, filters: string,
    orderby: { column: string; isAscending: boolean }, props: ISonaDeviationProps): Promise<IParameterType[]>;
}

export default function ParameterTypeOps() {
    const spCrudOps = SPCRUDOPS();

    const getAllParameterType = async (props: ISonaDeviationProps): Promise<IParameterType[]> => {
        return await (await spCrudOps).getData("Parameter", "Id,ParameterType,SubParameter", "", ""
            , { column: 'Id', isAscending: false }, props).then(results => {
                var output: Array<IParameterType> = new Array<IParameterType>();
                results.map((item: any) => {
                    output.push({
                        Id: item.Id,
                        ParameterType: item.ParameterType,
                        SubParameter: item.SubParameter
                    });
                });
                return output;
            });
    };

    const getParameterTypeById = async (Id: string | number, props: ISonaDeviationProps): Promise<IParameterType> => {
        return await (await spCrudOps).getData("Parameter", "Id,ParameterType,SubParameter", "", "ID eq " + Id + ""
            , { column: 'Id', isAscending: false }, props).then(results => {
                var output: Array<IParameterType> = new Array<IParameterType>();
                results.map((item: any) => {
                    output.push({
                        Id: item.Id,
                        ParameterType: item.ParameterType,
                        SubParameter: item.SubParameter
                    });
                });
                return output[0];
            });
    };

    const getParameterType = async (columnsToRetrieve: string, columnsToExpand: string, filters: string, orderby: { column: string; isAscending: boolean; }, p0: number, props: ISonaDeviationProps): Promise<IParameterType[]> => {
        return await (await spCrudOps).getData("Parameter", columnsToRetrieve, columnsToExpand, filters
            , orderby, props).then(results => {
                var output: Array<IParameterType> = new Array<IParameterType>();
                results.map((item: any) => {
                    output.push({
                        Id: item.Id,
                        ParameterType: item.ParameterType,
                        SubParameter: item.SubParameter
                    });
                });
                return output;
            });
    };

    const getTopParameterType = async (columnsToRetrieve: string, columnsToExpand: string, filters: string
            , orderby: { column: string, isAscending: boolean }, top: number, props: ISonaDeviationProps): Promise<IParameterType[]> => {
            return await (await spCrudOps).getTopData("Parameter", columnsToRetrieve, columnsToExpand, filters
                , orderby, top, props).then(results => {
                    var output: Array<IParameterType> = new Array<IParameterType>();
                    results.map((item: any) => {
                        output.push({
                            Id: item.Id,
                            PlantNameId: item.PlantName?.ID ?? null,
                            PlantName: item.PlantName?.PlantName || "",
                            ParameterType: item.ParameterType,
                            SubParameter: item.SubParameter
                        });
                    });
                    return output;
                });
        };
    return {
        getAllParameterType,
        getParameterTypeById,
        getParameterType,
        getTopParameterType
    };

}