
import { ISonaDeviationProps } from '../../components/ISonaDeviationProps';
import SPCRUDOPS from '../DAL/spcrudops';
import { IUserProfile } from '../../service/INTERFACE/IUserProfile';

export interface UserProfileOps {
    getLoggUserProfile(props: ISonaDeviationProps): Promise<IUserProfile>;
}

export default function LoggUserProfileOps(): UserProfileOps {
    const spCrudOps = SPCRUDOPS();

    const getLoggUserProfile = async (props: ISonaDeviationProps): Promise<IUserProfile> => {
        const result = await (await spCrudOps).currentProfile(props);

        // Debug log to inspect all properties
        //console.log("UserProfilePropertiesRS:", result.UserProfileProperties);

        const userProfileProperties = result.UserProfileProperties ?? [];

        // Attempt to find the location using a likely key
        const locationValue = userProfileProperties.find(
            prop => prop.Key === 'Office' // Change this key if needed
        )?.Value || "Location not found";

        return {
            AccountName: result.AccountName,
            UserProfileProperties: userProfileProperties,
            Location: locationValue
        };
    };

    return {
        getLoggUserProfile
    };
}
