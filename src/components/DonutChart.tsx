interface DonutChartProps {
    data: {
        value: number;
        label: string;
        color: string;
    }[];
    size?: number;
    thickness?: number;
}

const DonutChart: React.FC<DonutChartProps> = ({
    data,
    size = 200,
    thickness = 40,
    width = 100,
    height = 100
}) => {
    const radius = (size - thickness) / 2;
    const circumference = 2 * Math.PI * radius;
    const total = data.reduce((sum, item) => sum + item.value, 0);
    let offset = 0;     
    return (
        <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`}>
            <g transform={`translate(${size / 2}, ${size / 2})`}>
            
