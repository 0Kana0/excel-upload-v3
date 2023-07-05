import axios from 'axios';

export const AddVehicleBookingFromExcel = async (data) => {
	return await axios.post("http://192.168.0.145:8080/api/postvehiclebookingstatusfromexcel", data)
}